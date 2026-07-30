import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { AadHttpClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './HtmlFileRendererWebPart.module.scss';
import * as strings from 'HtmlFileRendererWebPartStrings';

export interface IHtmlFileRendererWebPartProps {
  htmlFileUrl: string;
  htmlCode: string;
  htmlCodeChunkCount?: number; // Number of chunks if content was split
  fullScreen: boolean;
  iframeHeight: string;
  lockDown: boolean;
  dataverseEnvironmentUrl: string;
  // Dynamic chunk properties: htmlCode_0, htmlCode_1, etc.
  [key: string]: string | number | boolean | undefined;
}

type JQueryCollectionLike = {
  length: number;
  click: () => void;
};

type WindowWithJQuery = Window & {
  $?: (selector: string) => JQueryCollectionLike;
  jQuery?: (selector: string) => JQueryCollectionLike;
};

type DataverseRequestInit = {
  method?: string;
  headers?: { [key: string]: string };
  body?: string;
};

type DataverseRequestMessage = {
  type?: string;
  requestId?: string;
  url?: string;
  init?: DataverseRequestInit;
};

type DataverseResponseMessage = {
  type: string;
  requestId: string;
  ok: boolean;
  status: number;
  statusText: string;
  headers: { [key: string]: string };
  body: string;
  error?: string;
};

export default class HtmlFileRendererWebPart extends BaseClientSideWebPart<IHtmlFileRendererWebPartProps> {

  // Chunking configuration: 250KB chunks × 120 max = ~30MB max
  private static readonly CHUNK_SIZE = 250000; // 250KB per chunk
  private static readonly MAX_CHUNKS = 120; // 120 chunks max
  private static readonly DESTRUCTIVE_OPERATION_WINDOW_MS = 60000;
  private static readonly DESTRUCTIVE_OPERATION_LIMIT = 5;
  private static readonly EDIT_PANEL_HOTKEY_MESSAGE = 'firepit:open-edit-panel';
  private static readonly DATAVERSE_REQUEST_MESSAGE = 'firepit:dataverse-request';
  private static readonly DATAVERSE_RESPONSE_MESSAGE = 'firepit:dataverse-response';
  private static readonly DESTRUCTIVE_COMMAND_APPROVED_MESSAGE = 'firepit:destructive-command-approved';

  // State for fetch-based rendering
  private _htmlContent: string = '';
  private _isLoading: boolean = false;
  private _error: string = '';
  private _destructiveCommandsApproved: boolean = false;
  private readonly _dataverseMutationTimestamps: { [targetKey: string]: number[] } = {};

  private readonly _handleEditPanelHotkey = (event: KeyboardEvent): void => {
    if (!this._isEditPanelHotkey(event) || event.repeat) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    this._openSharePointEditPanel();
  };

  private readonly _handleEditPanelHotkeyMessage = (event: MessageEvent): void => {
    if (!event.data || typeof event.data !== 'object') {
      return;
    }

    const data = event.data as { type?: string };
    if (data.type !== HtmlFileRendererWebPart.EDIT_PANEL_HOTKEY_MESSAGE) {
      return;
    }

    const iframe = this.domElement.querySelector<HTMLIFrameElement>('#htmlContentFrame');
    if (!iframe || event.source !== iframe.contentWindow) {
      return;
    }

    this._openSharePointEditPanel();
  };

  private readonly _handleDataverseMessage = (event: MessageEvent): void => {
    if (!event.data || typeof event.data !== 'object') {
      return;
    }

    const data = event.data as DataverseRequestMessage;
    if (data.type !== HtmlFileRendererWebPart.DATAVERSE_REQUEST_MESSAGE) {
      return;
    }

    const iframe = this.domElement.querySelector<HTMLIFrameElement>('#htmlContentFrame');
    const iframeWindow = iframe?.contentWindow;
    if (!iframeWindow || event.source !== iframeWindow || !data.requestId) {
      return;
    }

    this._handleDataverseRequest(data, iframeWindow).catch((error) => {
      this._postDataverseResponse(iframeWindow, {
        type: HtmlFileRendererWebPart.DATAVERSE_RESPONSE_MESSAGE,
        requestId: data.requestId || '',
        ok: false,
        status: 0,
        statusText: '',
        headers: {},
        body: '',
        error: error instanceof Error ? error.message : 'Dataverse request failed.'
      });
    });
  };

  private readonly _handleDestructiveCommandApprovalMessage = (event: MessageEvent): void => {
    if (!event.data || typeof event.data !== 'object') {
      return;
    }

    const data = event.data as { type?: string };
    if (data.type !== HtmlFileRendererWebPart.DESTRUCTIVE_COMMAND_APPROVED_MESSAGE) {
      return;
    }

    const iframe = this.domElement.querySelector<HTMLIFrameElement>('#htmlContentFrame');
    if (!iframe || event.source !== iframe.contentWindow) {
      return;
    }

    this._destructiveCommandsApproved = true;
  };

  public render(): void {
    // Check for content: URL, inline code, or chunked code
    const hasContent = this.properties.htmlFileUrl || this.properties.htmlCode || this._hasChunkedContent();
    if (!hasContent) {
      this.domElement.innerHTML = `
        <div class="${styles.htmlFileRenderer}">
          <div class="${styles.placeholder}">
            <div class="${styles.placeholderIcon}">
              <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <line x1="10" y1="9" x2="8" y2="9"></line>
              </svg>
            </div>
            <h2>${strings.PlaceholderTitle}</h2>
            <p>${strings.PlaceholderDescription}</p>
          </div>
        </div>`;
      return;
    }

    // Show loading state
    if (this._isLoading) {
      this.domElement.innerHTML = `
        <div class="${styles.htmlFileRenderer}">
          <div class="${styles.loading}">
            <div class="${styles.spinner}"></div>
            <p>${strings.LoadingMessage}</p>
          </div>
        </div>`;
      return;
    }

    // Show error state
    if (this._error) {
      this.domElement.innerHTML = `
        <div class="${styles.htmlFileRenderer}">
          <div class="${styles.error}">
            <div class="${styles.errorIcon}">⚠️</div>
            <h3>${strings.ErrorTitle}</h3>
            <p>${this._error}</p>
            <button class="${styles.retryButton}" id="retryButton">${strings.RetryButton}</button>
          </div>
        </div>`;

      // Attach retry handler
      const retryButton = this.domElement.querySelector('#retryButton');
      if (retryButton) {
        retryButton.addEventListener('click', () => this._loadHtmlFile());
      }
      return;
    }

    // Render the iframe with loaded content
    const height = this.properties.iframeHeight || '600px';
    const sandboxAttr = this._getSandboxAttribute();

    // Check for full screen mode
    let containerClass = styles.htmlFileRenderer;
    // Only enable full screen if configured AND we are NOT in edit mode
    // forcing full screen in edit mode makes the property pane inaccessible
    if (this.properties.fullScreen && this.displayMode !== DisplayMode.Edit) {
      containerClass = `${styles.htmlFileRenderer} ${styles.fullScreen}`;
    }

    // Build container without iframe in innerHTML to avoid MCAS js-wrapper
    // intercepting innerHTML and re-executing scripts inside the iframe
    this.domElement.innerHTML = `
      <div class="${containerClass}">
        <div class="${styles.iframeContainer}" id="htmlFrameContainer">
        </div>
      </div>`;

    const container = this.domElement.querySelector('#htmlFrameContainer');
    if (container && this._htmlContent) {
      // Create iframe programmatically and load via blob URL.
      // blob: URLs bypass MCAS js-wrapper interception of srcdoc content,
      // which otherwise causes scripts to be double-executed ("Identifier X
      // has already been declared" errors for every const/let in the app).
      const iframe = document.createElement('iframe');
      iframe.id = 'htmlContentFrame';
      iframe.className = styles.iframe;
      iframe.setAttribute('sandbox', sandboxAttr);
      iframe.style.height = height;
      iframe.title = 'HTML Content';

      const blob = new Blob([this._htmlContent], { type: 'text/html' });
      const blobUrl = URL.createObjectURL(blob);

      // Set src before appending to DOM so the browser loads blob content
      // directly, before any MutationObserver-based interception can fire
      iframe.src = blobUrl;

      // Revoke blob URL after load to prevent memory leaks
      iframe.addEventListener('load', () => {
        URL.revokeObjectURL(blobUrl);
        if (this._destructiveCommandsApproved) {
          this._notifyIframeDestructiveCommandApproval(iframe.contentWindow);
        }
      }, { once: true });

      container.appendChild(iframe);
    }
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    window.addEventListener('keydown', this._handleEditPanelHotkey, true);
    window.addEventListener('message', this._handleEditPanelHotkeyMessage);
    window.addEventListener('message', this._handleDataverseMessage);
    window.addEventListener('message', this._handleDestructiveCommandApprovalMessage);

    // Load HTML file if URL is configured and no code is provided
    if (this.properties.htmlFileUrl && !this.properties.htmlCode && !this._hasChunkedContent()) {
      await this._loadHtmlFile();
    } else {
      // Try to get content from chunks first, then fall back to single property
      const fullContent = this._reassembleChunks() || this.properties.htmlCode || '';
      if (fullContent) {
        this._htmlContent = this._processHtmlCode(fullContent);
        this.render();
      }
    }
  }

  protected onDispose(): void {
    window.removeEventListener('keydown', this._handleEditPanelHotkey, true);
    window.removeEventListener('message', this._handleEditPanelHotkeyMessage);
    window.removeEventListener('message', this._handleDataverseMessage);
    window.removeEventListener('message', this._handleDestructiveCommandApprovalMessage);
  }

  /**
   * Called when a property is changed in the property pane
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'htmlFileUrl' && newValue !== oldValue) {
      const normalizedUrl = this._normalizeHtmlFileUrl(newValue as string);
      if (normalizedUrl !== newValue) {
        this.properties.htmlFileUrl = normalizedUrl;
      }

      if (!this.properties.htmlCode) {
        this._loadHtmlFile().catch((err) => {
          this._error = err instanceof Error ? err.message : 'Failed to load HTML file';
          this.render();
        });
      }
    }

    if (propertyPath === 'htmlCode' && newValue !== oldValue) {
      const content = newValue as string;

      // Check if content needs to be chunked
      if (content && content.length > HtmlFileRendererWebPart.CHUNK_SIZE) {
        this._saveChunkedContent(content);
      } else {
        // Clear any existing chunks when using single property
        this._clearChunks();
        this.properties.htmlCodeChunkCount = undefined;
      }

      this._htmlContent = this._processHtmlCode(content);
      this.render();
    }

    if (propertyPath === 'dataverseEnvironmentUrl' && newValue !== oldValue) {
      const normalizedUrl = this._normalizeDataverseEnvironmentUrl(newValue as string);
      if (normalizedUrl !== newValue) {
        this.properties.dataverseEnvironmentUrl = normalizedUrl;
      }

      if (this.properties.htmlFileUrl && !this.properties.htmlCode && !this._hasChunkedContent()) {
        this._loadHtmlFile().catch((err) => {
          this._error = err instanceof Error ? err.message : 'Failed to load HTML file';
          this.render();
        });
        return;
      }

      const fullContent = this._reassembleChunks() || this.properties.htmlCode || '';
      if (fullContent) {
        this._htmlContent = this._processHtmlCode(fullContent);
        this.render();
      }
    }
  }

  /**
   * Load the HTML file from SharePoint
   */
  private async _loadHtmlFile(): Promise<void> {
    if (!this.properties.htmlFileUrl) {
      return;
    }

    this._isLoading = true;
    this._error = '';
    this.render();

    try {
      const fileUrl = this._resolveFileUrl(this.properties.htmlFileUrl);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        fileUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`${strings.HttpErrorPrefix} ${response.status}: ${response.statusText}`);
      }

      const rawHtml = await response.text();
      // Clean the HTML to remove MCAS-injected scripts that cause CORS issues
      const cleanedHtml = this._cleanHtmlContent(rawHtml);
      // Process HTML to inject SharePoint context and CSP nonces
      this._htmlContent = this._processHtmlCode(cleanedHtml);
      this._isLoading = false;
      this._error = '';
    } catch (error) {
      this._isLoading = false;
      this._error = error instanceof Error ? error.message : strings.UnknownError;
    }

    this.render();
  }

  /**
   * Clean HTML content by removing MCAS (Microsoft Cloud App Security) injected scripts.
   * "Nuclear Option": Since the target file is a self-contained app (Scheduler.html),
   * we remove ALL external scripts and styles that might be injected by MCAS/SharePoint.
   */
  private _cleanHtmlContent(html: string): string {
    let cleaned = html;

    // Clean inline attributes (onclick, etc) that might be wrapped
    cleaned = cleaned.replace(/__WRAPPED_\w+/g, 'void(0)');

    // Remove any meta tags added by MCAS
    cleaned = cleaned.replace(/<meta[^>]*(?:mcas|McasCtx|McasTsid)[^>]*\/?>/gi, '');

    // Remove MCAS js-wrapper script tags injected into fetched HTML
    cleaned = cleaned.replace(/<script[^>]*src="[^"]*js-wrapper[^"]*"[^>]*><\/script>/gi, '');

    // Strip MCAS query string parameters from src/href URLs to avoid redirect loops
    cleaned = cleaned.replace(/((?:src|href)\s*=\s*["'][^"']*?)\?McasCtx=[^"'&]*(?:&McasTsid=[^"'&]*)?(?:&McasUserAuth=[^"']*)?/gi, '$1');

    return cleaned;
  }

  /**
   * Check if content is stored in chunks
   */
  private _hasChunkedContent(): boolean {
    return typeof this.properties.htmlCodeChunkCount === 'number' &&
      this.properties.htmlCodeChunkCount > 0 &&
      // eslint-disable-next-line dot-notation
      typeof this.properties['htmlCode_0'] === 'string';
  }

  /**
   * Reassemble chunked content from multiple properties
   */
  private _reassembleChunks(): string {
    const chunkCount = this.properties.htmlCodeChunkCount;
    if (!chunkCount || chunkCount <= 0) {
      return '';
    }

    let content = '';
    for (let i = 0; i < chunkCount; i++) {
      const chunk = this.properties[`htmlCode_${i}`];
      if (typeof chunk === 'string') {
        content += chunk;
      }
    }
    return content;
  }

  /**
   * Save content by splitting into chunks
   */
  private _saveChunkedContent(content: string): void {
    // Clear existing chunks first
    this._clearChunks();

    // Split content into chunks
    const chunks: string[] = [];
    for (let i = 0; i < content.length; i += HtmlFileRendererWebPart.CHUNK_SIZE) {
      chunks.push(content.substring(i, i + HtmlFileRendererWebPart.CHUNK_SIZE));
    }

    // Check if we exceed max chunks
    if (chunks.length > HtmlFileRendererWebPart.MAX_CHUNKS) {
      console.warn(`Content requires ${chunks.length} chunks, but max is ${HtmlFileRendererWebPart.MAX_CHUNKS}. Content will be truncated.`);
    }

    // Save chunks to properties
    const chunkCount = Math.min(chunks.length, HtmlFileRendererWebPart.MAX_CHUNKS);
    for (let i = 0; i < chunkCount; i++) {
      this.properties[`htmlCode_${i}`] = chunks[i];
    }
    this.properties.htmlCodeChunkCount = chunkCount;

    // Clear the main htmlCode property to avoid duplication
    // But keep it for display in the property pane
    // Actually, we'll keep the full content in htmlCode for the textarea
    // The chunks are for storage only
  }

  /**
   * Clear all chunk properties
   */
  private _clearChunks(): void {
    // Clear chunk count
    const existingCount = this.properties.htmlCodeChunkCount || 0;

    // Remove each chunk property
    for (let i = 0; i < existingCount; i++) {
      delete this.properties[`htmlCode_${i}`];
    }

    this.properties.htmlCodeChunkCount = undefined;
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Resolve the file URL to a full SharePoint URL
   */
  private _resolveFileUrl(url: string): string {
    const normalizedUrl = this._normalizeHtmlFileUrl(url);

    // If it's already a full URL, return as-is
    if (normalizedUrl.startsWith('http://') || normalizedUrl.startsWith('https://')) {
      return normalizedUrl;
    }

    // If it starts with a slash, it's site-relative
    if (normalizedUrl.startsWith('/')) {
      return `${this.context.pageContext.web.absoluteUrl}${normalizedUrl}`;
    }

    // Otherwise, assume it's relative to the site
    return `${this.context.pageContext.web.absoluteUrl}/${normalizedUrl}`;
  }

  private _normalizeHtmlFileUrl(url: string): string {
    const trimmedUrl = (url || '').trim();
    const htmlExtensionIndex = trimmedUrl.toLowerCase().indexOf('.html');
    if (htmlExtensionIndex === -1) {
      return trimmedUrl;
    }

    return trimmedUrl.substring(0, htmlExtensionIndex + '.html'.length);
  }

  private _normalizeDataverseEnvironmentUrl(url: string): string {
    const trimmedUrl = (url || '').trim();
    if (!trimmedUrl) {
      return '';
    }

    try {
      const parsed = new URL(trimmedUrl);
      parsed.hash = '';
      parsed.search = '';
      return parsed.origin;
    } catch {
      return trimmedUrl.replace(/\/+$/, '');
    }
  }

  private _getDataverseOrigin(): string | undefined {
    const configuredUrl = this._normalizeDataverseEnvironmentUrl(this.properties.dataverseEnvironmentUrl || '');
    if (!configuredUrl) {
      return undefined;
    }

    try {
      const parsed = new URL(configuredUrl);
      if (parsed.protocol !== 'https:') {
        return undefined;
      }

      return parsed.origin;
    } catch {
      return undefined;
    }
  }

  private _resolveDataverseRequestUrl(rawUrl: string | undefined): string {
    if (!rawUrl) {
      throw new Error('Dataverse request URL is required.');
    }

    const dataverseOrigin = this._getDataverseOrigin();
    if (!dataverseOrigin) {
      throw new Error('Dataverse environment URL is not configured.');
    }

    const targetUrl = new URL(rawUrl, dataverseOrigin);
    if (targetUrl.origin !== dataverseOrigin) {
      throw new Error('Dataverse request URL must stay within the configured Dataverse environment.');
    }

    return targetUrl.toString();
  }

  private async _handleDataverseRequest(message: DataverseRequestMessage, targetWindow: Window): Promise<void> {
    if (this.properties.lockDown) {
      throw new Error('Dataverse requests are disabled while Lock Down Mode is on.');
    }

    const dataverseOrigin = this._getDataverseOrigin();
    if (!dataverseOrigin) {
      throw new Error('Dataverse environment URL is not configured.');
    }

    const requestUrl = this._resolveDataverseRequestUrl(message.url);
    const requestInit = this._buildDataverseRequestInit(message.init);
    if (!this._confirmDataverseRequestSafety(requestUrl, requestInit)) {
      throw new Error('Blocked by Firepit: command not confirmed.');
    }

    const client = await this.context.aadHttpClientFactory.getClient(dataverseOrigin);
    const response = await client.fetch(requestUrl, AadHttpClient.configurations.v1, requestInit);
    const headers: { [key: string]: string } = {};
    response.headers.forEach((value, key) => {
      headers[key] = value;
    });

    this._postDataverseResponse(targetWindow, {
      type: HtmlFileRendererWebPart.DATAVERSE_RESPONSE_MESSAGE,
      requestId: message.requestId || '',
      ok: response.ok,
      status: response.status,
      statusText: response.statusText,
      headers,
      body: await response.text()
    });
  }

  private _buildDataverseRequestInit(init: DataverseRequestInit | undefined): RequestInit {
    const method = (init?.method || 'GET').toUpperCase();
    const requestInit: RequestInit = {
      method,
      headers: this._buildDataverseHeaders(init?.headers)
    };

    if (init?.body !== undefined && method !== 'GET' && method !== 'HEAD') {
      requestInit.body = init.body;
    }

    return requestInit;
  }

  private _buildDataverseHeaders(inputHeaders: { [key: string]: string } | undefined): { [key: string]: string } {
    const headers: { [key: string]: string } = {
      Accept: 'application/json',
      'Content-Type': 'application/json; charset=utf-8',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0'
    };

    Object.keys(inputHeaders || {}).forEach((key) => {
      const normalizedKey = key.toLowerCase();
      if ([
        'authorization',
        'cookie',
        'host',
        'origin',
        'referer',
        'content-length'
      ].indexOf(normalizedKey) !== -1) {
        return;
      }

      headers[key] = String(inputHeaders?.[key] || '');
    });

    return headers;
  }

  private _postDataverseResponse(targetWindow: Window, message: DataverseResponseMessage): void {
    targetWindow.postMessage(message, '*');
  }

  private _notifyIframeDestructiveCommandApproval(targetWindow: Window | null | undefined): void {
    if (!targetWindow) {
      return;
    }

    targetWindow.postMessage({
      type: HtmlFileRendererWebPart.DESTRUCTIVE_COMMAND_APPROVED_MESSAGE
    }, '*');
  }

  private _confirmDataverseRequestSafety(requestUrl: string, requestInit: RequestInit): boolean {
    if (this._destructiveCommandsApproved) {
      return true;
    }

    const operations = this._analyzeDataverseOperations(requestUrl, requestInit);
    if (operations.length === 0) {
      return true;
    }

    const reason = this._getDataverseConfirmationReason(operations);
    if (!reason) {
      this._recordDataverseItemMutations(operations);
      return true;
    }

    const approved = window.confirm(`Firepit confirmation required\n\n${reason}\n\nAllow this command to be sent?`);
    if (approved) {
      this._destructiveCommandsApproved = true;
      this._recordDataverseItemMutations(operations);
      const iframe = this.domElement.querySelector<HTMLIFrameElement>('#htmlContentFrame');
      this._notifyIframeDestructiveCommandApproval(iframe?.contentWindow);
      return true;
    }

    return false;
  }

  private _analyzeDataverseOperations(requestUrl: string, requestInit: RequestInit): Array<{ kind: string; targetKey: string; targetLabel: string }> {
    let parsedUrl: URL;
    try {
      parsedUrl = new URL(requestUrl);
    } catch {
      return [];
    }

    const path = this._safeDecodeUriComponent(parsedUrl.pathname.toLowerCase());
    const body = typeof requestInit.body === 'string' ? requestInit.body : '';
    const method = this._getEffectiveDataverseMethod(requestInit);

    if (/\/api\/data\/v[0-9.]+\/\$batch$/i.test(path)) {
      return this._analyzeDataverseBatchOperations(body, parsedUrl.origin);
    }

    const target = this._getDataverseOperationTarget(path);
    const operations: Array<{ kind: string; targetKey: string; targetLabel: string }> = [];
    if (method === 'DELETE' && /\/entitydefinitions\(/i.test(path)) {
      operations.push({ kind: 'full-delete', targetKey: target.key, targetLabel: target.label });
    } else if (/^(DELETE|MERGE|PATCH|PUT)$/.test(method) && /\/api\/data\/v[0-9.]+\/[^/?#(]+\([^)]*\)/i.test(path)) {
      operations.push({ kind: 'item', targetKey: target.key, targetLabel: target.label });
    }

    return operations;
  }

  private _analyzeDataverseBatchOperations(body: string, origin: string): Array<{ kind: string; targetKey: string; targetLabel: string }> {
    if (!body) {
      return this._buildUnknownDataverseBatchOperations();
    }

    const operations: Array<{ kind: string; targetKey: string; targetLabel: string }> = [];
    const direct = /(?:^|\r?\n)(DELETE|PATCH|MERGE|PUT)\s+([^\s]+)\s+HTTP\/1\.[01]/gi;
    let match: RegExpExecArray | null;
    while ((match = direct.exec(body)) !== null) {
      operations.push(...this._analyzeDataverseOperations(new URL(match[2], origin).toString(), { method: match[1] }));
    }

    const override = /(?:^|\r?\n)POST\s+([^\s]+)\s+HTTP\/1\.[01][\s\S]*?\r?\nX-HTTP-Method(?:-Override)?:\s*(DELETE|MERGE|PATCH|PUT)/gi;
    while ((match = override.exec(body)) !== null) {
      operations.push(...this._analyzeDataverseOperations(new URL(match[1], origin).toString(), { method: match[2] }));
    }

    return operations;
  }

  private _buildUnknownDataverseBatchOperations(): Array<{ kind: string; targetKey: string; targetLabel: string }> {
    const operations: Array<{ kind: string; targetKey: string; targetLabel: string }> = [];
    for (let i = 0; i <= HtmlFileRendererWebPart.DESTRUCTIVE_OPERATION_LIMIT; i++) {
      operations.push({ kind: 'item', targetKey: 'dataverse:batch', targetLabel: 'a Dataverse batch request' });
    }

    return operations;
  }

  private _getDataverseOperationTarget(path: string): { key: string; label: string } {
    const tableDefinitionMatch = path.match(/\/entitydefinitions\(logicalname=['"]?([^'")]+)['"]?\)/i);
    if (tableDefinitionMatch) {
      const logicalName = tableDefinitionMatch[1];
      return {
        key: `dataverse:${logicalName.toLowerCase()}`,
        label: `Dataverse table ${logicalName}`
      };
    }

    const match = path.match(/\/api\/data\/v[0-9.]+\/([^/?#(]+)/i);
    const label = match?.[1] || 'Dataverse table';
    return {
      key: `dataverse:${label.toLowerCase()}`,
      label
    };
  }

  private _getEffectiveDataverseMethod(requestInit: RequestInit): string {
    const headers = requestInit.headers as { [key: string]: string } | undefined;
    const method = (requestInit.method || 'GET').toUpperCase();
    const override = (headers?.['X-HTTP-Method'] || headers?.['x-http-method'] ||
      headers?.['X-HTTP-Method-Override'] || headers?.['x-http-method-override'] || '').toUpperCase();
    if (method === 'POST' && /^(DELETE|MERGE|PATCH|PUT)$/.test(override)) {
      return override;
    }

    return method;
  }

  private _getDataverseConfirmationReason(operations: Array<{ kind: string; targetKey: string; targetLabel: string }>): string {
    const fullDelete = operations.find((operation) => operation.kind === 'full-delete');
    if (fullDelete) {
      return `This embedded app is trying to delete the entire ${fullDelete.targetLabel}.`;
    }

    const now = Date.now();
    const counts: { [targetKey: string]: number } = {};
    const labels: { [targetKey: string]: string } = {};
    operations.forEach((operation) => {
      if (operation.kind !== 'item') {
        return;
      }

      counts[operation.targetKey] = (counts[operation.targetKey] || 0) + 1;
      labels[operation.targetKey] = operation.targetLabel;
    });

    for (const targetKey of Object.keys(counts)) {
      const currentCount = this._pruneDataverseMutationHistory(targetKey, now).length;
      if (currentCount + counts[targetKey] > HtmlFileRendererWebPart.DESTRUCTIVE_OPERATION_LIMIT) {
        return `This embedded app is trying to edit or delete ${currentCount + counts[targetKey]} items from ${labels[targetKey]} within one minute.`;
      }
    }

    return '';
  }

  private _recordDataverseItemMutations(operations: Array<{ kind: string; targetKey: string }>): void {
    const now = Date.now();
    operations.forEach((operation) => {
      if (operation.kind !== 'item') {
        return;
      }

      this._pruneDataverseMutationHistory(operation.targetKey, now).push(now);
    });
  }

  private _pruneDataverseMutationHistory(targetKey: string, now: number): number[] {
    const existingHistory = this._dataverseMutationTimestamps[targetKey] || [];
    this._dataverseMutationTimestamps[targetKey] = existingHistory.filter((timestamp) => {
      return now - timestamp < HtmlFileRendererWebPart.DESTRUCTIVE_OPERATION_WINDOW_MS;
    });

    return this._dataverseMutationTimestamps[targetKey];
  }

  private _safeDecodeUriComponent(value: string): string {
    try {
      return decodeURIComponent(value);
    } catch {
      return value;
    }
  }

  /**
   * Inject SharePoint context and security hardening into the HTML code.
   *
   * Security layers ported from the Forge compiler:
   * 1. CSP meta tag restricts connect-src to SharePoint origin, or none in lock-down
   * 2. Additional security meta tags
   * 3. Link-hint guard neutralizes link hints that can trigger outbound requests
   * 4. Query-param guard strips query params from navigation URLs
   * 5. Network guard blocks non-SharePoint fetch, XHR, WebSocket, EventSource, and sendBeacon
   */
  private _processHtmlCode(code: string): string {
    if (!code) return '';

    const webUrl = this.context.pageContext.web.absoluteUrl;
    const userEmail = this.context.pageContext.user.email;
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const cspNonce = this._getHostCspNonce();
    const nonceAttr = cspNonce ? ` nonce="${this._escapeHtmlAttribute(cspNonce)}"` : '';

    let spOrigin = '';
    try {
      spOrigin = new URL(webUrl).origin;
    } catch {
      // Leave empty when the page context URL cannot be parsed.
    }

    const mcasWrappedSharePointOrigin = this._getMcasWrappedSharePointOrigin(spOrigin);
    const dataverseOrigin = this._getDataverseOrigin();

    const contextScript = `
      <script${nonceAttr}>
        window._spPageContextInfo = {
          webAbsoluteUrl: "${webUrl}",
          siteAbsoluteUrl: "${this.context.pageContext.site.absoluteUrl}",
          webId: "${webId}",
          siteId: "${siteId}",
          userEmail: "${userEmail}",
          userId: "${this.context.pageContext.user.loginName}"
        };
        window.__firepitCspNonce = ${JSON.stringify(cspNonce ?? '')};
        window.__firepitDataverse = {
          enabled: ${JSON.stringify(Boolean(dataverseOrigin) && !this.properties.lockDown)},
          environmentUrl: ${JSON.stringify(dataverseOrigin || '')}
        };
      </script>
    `;

    const mcasOrigin = 'https://mcas-proxyweb.mcas-gov.us';
    const mcasCdnOrigin = 'https://inline.cdn.mcas-gov.us';
    const connectSrcOrigins = [
      "'self'",
      spOrigin,
      mcasWrappedSharePointOrigin,
      mcasOrigin,
      mcasCdnOrigin,
      dataverseOrigin
    ].filter(Boolean).join(' ');
    const connectSrc = this.properties.lockDown
      ? "connect-src 'none'"
      : `connect-src ${connectSrcOrigins}`;
    const frameSrcOrigins = [
      spOrigin,
      mcasWrappedSharePointOrigin
    ].filter(Boolean).join(' ');
    const frameSrc = this.properties.lockDown
      ? "frame-src 'none'"
      : (frameSrcOrigins ? `frame-src ${frameSrcOrigins}` : "frame-src 'none'");
    const cspDirectives = [
      "default-src 'none'",
      "script-src 'unsafe-inline' 'unsafe-eval'",
      "style-src 'unsafe-inline'",
      'worker-src blob:',
      connectSrc,
      "img-src data: blob: 'self'",
      'font-src data:',
      "media-src 'none'",
      "manifest-src 'none'",
      "form-action 'none'",
      frameSrc,
      "object-src 'none'"
    ].join('; ') + ';';
    const cspMeta = `<meta http-equiv="Content-Security-Policy" content="${cspDirectives}">`;

    const securityMetas = [
      cspMeta,
      '<meta http-equiv="x-dns-prefetch-control" content="off">',
      '<meta http-equiv="X-Content-Type-Options" content="nosniff">',
      '<meta http-equiv="X-XSS-Protection" content="1; mode=block">',
      '<meta http-equiv="Referrer-Policy" content="no-referrer">',
      '<meta http-equiv="Permissions-Policy" content="geolocation=(), payment=(), usb=(), magnetometer=(), gyroscope=(), accelerometer=(), ambient-light-sensor=(), autoplay=(), encrypted-media=(), fullscreen=(), picture-in-picture=(), screen-wake-lock=()">'
    ].join('\n    ');

    const linkHintGuard = `<script${nonceAttr}>${this._buildLinkHintGuardScript()}</script>`;
    const spSiteUrl = webUrl.replace(/\/+$/, '');
    const queryParamGuard = `<script${nonceAttr}>${this._buildQueryParamGuardScript(spOrigin, spSiteUrl, mcasWrappedSharePointOrigin)}</script>`;
    const editPanelHotkeyBridge = `<script${nonceAttr}>${this._buildEditPanelHotkeyBridgeScript()}</script>`;
    const destructiveCommandGuard = `<script${nonceAttr}>${this._buildDestructiveCommandGuardScript()}</script>`;
    const dataverseBridge = dataverseOrigin && !this.properties.lockDown
      ? `<script${nonceAttr}>${this._buildDataverseBridgeScript(dataverseOrigin)}</script>`
      : '';
    const networkGuard = this.properties.lockDown
      ? ''
      : `<script${nonceAttr}>${this._buildNetworkGuardScript(spOrigin, mcasWrappedSharePointOrigin, dataverseOrigin)}</script>`;

    const securityInjection = [
      '<!-- Firepit: Data-exfiltration prevention (ported from Forge) -->',
      securityMetas,
      linkHintGuard,
      queryParamGuard,
      editPanelHotkeyBridge,
      destructiveCommandGuard,
      dataverseBridge,
      networkGuard
    ].filter(Boolean).join('\n    ');

    const nonceReadyCode = cspNonce ? this._applyNonceToScriptTags(code, cspNonce) : code;
    return this._injectIntoHead(nonceReadyCode, contextScript + '\n    ' + securityInjection);
  }

  private _isEditPanelHotkey(event: KeyboardEvent): boolean {
    return event.ctrlKey &&
      event.altKey &&
      event.shiftKey &&
      !event.metaKey &&
      (event.code === 'KeyE' || event.key.toLowerCase() === 'e');
  }

  private _openSharePointEditPanel(): void {
    if (this._clickEditButtonWithJQuery()) {
      return;
    }

    const selectors = [
      '[role="menuitem"][name="Edit"]',
      '[name="Edit"]',
      'button[aria-label="Edit"]',
      '[role="button"][aria-label="Edit"]',
      'button[title="Edit"]',
      'a[title="Edit"]',
      '[data-automation-id="pageCommandBar"] [name="Edit"]',
      '[data-automationid="pageCommandBar"] [name="Edit"]'
    ];

    for (const selector of selectors) {
      const elements = Array.from(document.querySelectorAll<HTMLElement>(selector));
      const target = elements.find((element) => this._isVisibleClickableElement(element));
      if (target) {
        target.click();
        return;
      }
    }

    console.warn('Firepit edit-panel hotkey could not find a SharePoint Edit button.');
  }

  private _clickEditButtonWithJQuery(): boolean {
    const hostWindow = window as WindowWithJQuery;
    const jquery = hostWindow.$ || hostWindow.jQuery;
    if (typeof jquery !== 'function') {
      return false;
    }

    try {
      const editButton = jquery('[role="menuitem"][name="Edit"]');
      if (editButton.length > 0) {
        editButton.click();
        return true;
      }

      const fallbackEditButton = jquery('[name="Edit"]');
      if (fallbackEditButton.length > 0) {
        fallbackEditButton.click();
        return true;
      }
    } catch (error) {
      console.warn('Firepit edit-panel hotkey jQuery click failed:', error);
    }

    return false;
  }

  private _isVisibleClickableElement(element: HTMLElement): boolean {
    const button = element as HTMLButtonElement;
    if (button.disabled || element.getAttribute('aria-disabled') === 'true') {
      return false;
    }

    const rect = element.getBoundingClientRect();
    const style = window.getComputedStyle(element);
    return rect.width > 0 &&
      rect.height > 0 &&
      style.display !== 'none' &&
      style.visibility !== 'hidden';
  }

  private _buildEditPanelHotkeyBridgeScript(): string {
    const messageType = JSON.stringify(HtmlFileRendererWebPart.EDIT_PANEL_HOTKEY_MESSAGE);
    return [
      '(function(){',
      '  try {',
      '    var messageType = ' + messageType + ';',
      '    var isHotkey = function(e){',
      '      return !!e && e.ctrlKey && e.altKey && e.shiftKey && !e.metaKey && (e.code === "KeyE" || String(e.key || "").toLowerCase() === "e");',
      '    };',
      '    document.addEventListener("keydown",function(e){',
      '      if(!isHotkey(e) || e.repeat) return;',
      '      e.preventDefault();',
      '      e.stopPropagation();',
      '      try { window.parent.postMessage({ type: messageType }, "*"); } catch(err){}',
      '    },true);',
      '  } catch(e){ console.warn("Firepit edit-panel hotkey init failed:", e); }',
      '})();'
    ].join('');
  }

  private _buildDataverseBridgeScript(dataverseOrigin: string): string {
    const environmentUrl = JSON.stringify(dataverseOrigin);
    const requestType = JSON.stringify(HtmlFileRendererWebPart.DATAVERSE_REQUEST_MESSAGE);
    const responseType = JSON.stringify(HtmlFileRendererWebPart.DATAVERSE_RESPONSE_MESSAGE);

    return [
      '(function(){',
      '  try {',
      '    var ENVIRONMENT_URL = ' + environmentUrl + ';',
      '    var REQUEST_TYPE = ' + requestType + ';',
      '    var RESPONSE_TYPE = ' + responseType + ';',
      '    var sequence = 0;',
      '    var pending = {};',
      '    var toHeaderObject = function(headers){',
      '      var result = {};',
      '      if(!headers) return result;',
      '      try {',
      '        if(typeof Headers !== "undefined" && headers instanceof Headers){ headers.forEach(function(value,key){ result[key]=value; }); return result; }',
      '        if(Array.isArray(headers)){ headers.forEach(function(pair){ if(pair&&pair.length>=2) result[String(pair[0])]=String(pair[1]); }); return result; }',
      '        Object.keys(headers).forEach(function(key){ result[key]=String(headers[key]); });',
      '      } catch(e){}',
      '      return result;',
      '    };',
      '    var resolveRequest = function(message){',
      '      var item = pending[message.requestId];',
      '      if(!item) return;',
      '      delete pending[message.requestId];',
      '      clearTimeout(item.timer);',
      '      if(message.error){ item.reject(new Error(message.error)); return; }',
      '      var responseBody = [204,205,304].indexOf(message.status) !== -1 ? null : (message.body || "");',
      '      item.resolve(new Response(responseBody, { status: message.status || 200, statusText: message.statusText || "", headers: message.headers || {} }));',
      '    };',
      '    window.addEventListener("message",function(event){',
      '      var message = event.data;',
      '      if(!message || typeof message !== "object" || message.type !== RESPONSE_TYPE) return;',
      '      resolveRequest(message);',
      '    });',
      '    var firepitFetch = function(input, init){',
      '      init = init || {};',
      '      var url = "";',
      '      if(typeof input === "string" || input instanceof URL) url = String(input);',
      '      else if(input && typeof input.url === "string") url = input.url;',
      '      else return Promise.reject(new Error("Dataverse request URL is required."));',
      '      if(input && typeof Request !== "undefined" && input instanceof Request && !init.method) init.method = input.method;',
      '      var requestId = "dataverse-" + Date.now() + "-" + (++sequence);',
      '      var payload = {',
      '        type: REQUEST_TYPE,',
      '        requestId: requestId,',
      '        url: url,',
      '        init: {',
      '          method: init.method || "GET",',
      '          headers: toHeaderObject(init.headers),',
      '          body: init.body == null ? undefined : String(init.body)',
      '        }',
      '      };',
      '      return new Promise(function(resolve,reject){',
      '        pending[requestId] = { resolve: resolve, reject: reject, timer: setTimeout(function(){ delete pending[requestId]; reject(new Error("Dataverse request timed out.")); }, 120000) };',
      '        window.parent.postMessage(payload, "*");',
      '      });',
      '    };',
      '    var api = {',
      '      environmentUrl: ENVIRONMENT_URL,',
      '      fetch: firepitFetch,',
      '      request: function(url, init){',
      '        return firepitFetch(url, init).then(function(response){',
      '          if(!response.ok) throw new Error("Dataverse request failed: " + response.status + " " + response.statusText);',
      '          var contentType = response.headers.get("content-type") || "";',
      '          return contentType.indexOf("application/json") !== -1 ? response.json() : response.text();',
      '        });',
      '      }',
      '    };',
      '    window.firepit = window.firepit || {};',
      '    window.firepit.dataverse = api;',
      '    window.firepitDataverse = api;',
      '  } catch(e){ console.warn("Firepit Dataverse bridge init failed:",e); }',
      '})();'
    ].join('');
  }

  private _buildDestructiveCommandGuardScript(): string {
    const approvalMessageType = JSON.stringify(HtmlFileRendererWebPart.DESTRUCTIVE_COMMAND_APPROVED_MESSAGE);
    return [
      '(function(){',
      '  try {',
      '    var APPROVAL_MESSAGE_TYPE=' + approvalMessageType + ';',
      '    var WINDOW_MS=60000;',
      '    var LIMIT=5;',
      '    var history={};',
      '    var destructiveCommandsApproved=false;',
      '    window.addEventListener("message",function(event){ var message=event.data; if(message&&typeof message==="object"&&message.type===APPROVAL_MESSAGE_TYPE) destructiveCommandsApproved=true; });',
      '    var approveForSession=function(){ destructiveCommandsApproved=true; try{ window.parent.postMessage({ type: APPROVAL_MESSAGE_TYPE }, "*"); }catch(e){} };',
      '    var asString=function(v){ return v==null?"":String(v); };',
      '    var normalizePath=function(v){ var p=asString(v).toLowerCase(); try{ p=decodeURIComponent(p); }catch(e){} return p.replace(/\\/+/g,"/"); };',
      '    var headersToObject=function(headers){',
      '      var result={};',
      '      if(!headers) return result;',
      '      try {',
      '        if(typeof Headers!=="undefined"&&headers instanceof Headers){ headers.forEach(function(value,key){ result[String(key).toLowerCase()]=String(value); }); return result; }',
      '        if(Array.isArray(headers)){ headers.forEach(function(pair){ if(pair&&pair.length>=2) result[String(pair[0]).toLowerCase()]=String(pair[1]); }); return result; }',
      '        Object.keys(headers).forEach(function(key){ result[String(key).toLowerCase()]=String(headers[key]); });',
      '      } catch(e){}',
      '      return result;',
      '    };',
      '    var getEffectiveMethod=function(init){',
      '      init=init||{};',
      '      var headers=headersToObject(init.headers);',
      '      var method=asString(init.method||"GET").toUpperCase();',
      '      var override=asString(headers["x-http-method"]||headers["x-http-method-override"]).toUpperCase();',
      '      if(method==="POST"&&/^(DELETE|MERGE|PATCH|PUT)$/.test(override)) method=override;',
      '      return method;',
      '    };',
      '    var getBody=function(init){',
      '      if(!init||init.body==null) return "";',
      '      return typeof init.body==="string"?init.body:asString(init.body);',
      '    };',
      '    var makeOp=function(kind,action,targetKey,targetLabel){',
      '      return { kind:kind, action:action, targetKey:targetKey||"unknown", targetLabel:targetLabel||"the target list or table" };',
      '    };',
      '    var sharePointTarget=function(path){',
      '      var match=path.match(/\\/_api\\/web\\/lists\\/([^/?#]+(?:\\([^/?#]*\\))?)/i)||path.match(/\\/_api\\/web\\/getlist\\(([^)]*)\\)/i);',
      '      var label=match?match[1]:"SharePoint list";',
      '      label=label.replace(/^getbytitle\\(/i,"").replace(/^getbyid\\(/i,"").replace(/^guid/i,"").replace(/[()\\\'"]/g,"");',
      '      return { key:"sharepoint:"+label.toLowerCase(), label:label||"SharePoint list" };',
      '    };',
      '    var dataverseTarget=function(path){',
      '      var tableDefinitionMatch=path.match(/\\/entitydefinitions\\(logicalname=[\\\'"]?([^\\\'")]+)[\\\'"]?\\)/i);',
      '      if(tableDefinitionMatch){ return { key:"dataverse:"+tableDefinitionMatch[1].toLowerCase(), label:"Dataverse table "+tableDefinitionMatch[1] }; }',
      '      var match=path.match(/\\/api\\/data\\/v[0-9.]+\\/([^/?#(]+)/i);',
      '      var label=match?match[1]:"Dataverse table";',
      '      return { key:"dataverse:"+label.toLowerCase(), label:label||"Dataverse table" };',
      '    };',
      '    var analyzeSingle=function(url,init){',
      '      var operations=[];',
      '      var parsed; try{ parsed=new URL(asString(url),document.baseURI||location.href); }catch(e){ return operations; }',
      '      var path=normalizePath(parsed.pathname);',
      '      var method=getEffectiveMethod(init);',
      '      var body=getBody(init);',
      '      var isDelete=method==="DELETE";',
      '      var isEdit=/^(MERGE|PATCH|PUT)$/.test(method);',
      '      if(method==="POST"&&/validateupdatelistitem|systemupdate|update\\(/i.test(path)) isEdit=true;',
      '      if(method==="POST"&&/\\/recycle\\(\\)?$/i.test(path)) isDelete=true;',
      '      if(/\\/_api\\/(?:\\$batch|.*\\/\\$batch)$/i.test(path)||/\\/api\\/data\\/v[0-9.]+\\/\\$batch$/i.test(path)){',
      '        return analyzeBatch(body,parsed.origin);',
      '      }',
      '      if(path.indexOf("/_api/")!==-1&&path.indexOf("/web/lists")!==-1){',
      '        var spTarget=sharePointTarget(path);',
      '        var itemUrl=/\\/items?\\(|\\/items?\\/|\\/getitembyid\\(|validateupdatelistitem|\\/recycle\\(\\)?$/i.test(path);',
      '        if(isDelete&&!itemUrl) operations.push(makeOp("full-delete","delete",spTarget.key,spTarget.label));',
      '        else if((isDelete||isEdit)&&itemUrl) operations.push(makeOp("item",isDelete?"delete":"edit",spTarget.key,spTarget.label));',
      '      }',
      '      if(/\\/api\\/data\\/v[0-9.]+\\//i.test(path)){',
      '        var dvTarget=dataverseTarget(path);',
      '        if(isDelete&&/\\/entitydefinitions\\(/i.test(path)) operations.push(makeOp("full-delete","delete",dvTarget.key,dvTarget.label));',
      '        else if((isDelete||isEdit)&&/\\/api\\/data\\/v[0-9.]+\\/[^/?#(]+\\([^)]*\\)/i.test(path)) operations.push(makeOp("item",isDelete?"delete":"edit",dvTarget.key,dvTarget.label));',
      '      }',
      '      return operations;',
      '    };',
      '    var analyzeBatch=function(body,origin){',
      '      var operations=[];',
      '      if(!body||body==="[object FormData]"||body==="[object Blob]"||body==="[object ArrayBuffer]"){',
      '        operations.push(makeOp("item","edit/delete","batch","a batch request"));',
      '        operations.push(makeOp("item","edit/delete","batch","a batch request"));',
      '        operations.push(makeOp("item","edit/delete","batch","a batch request"));',
      '        operations.push(makeOp("item","edit/delete","batch","a batch request"));',
      '        operations.push(makeOp("item","edit/delete","batch","a batch request"));',
      '        operations.push(makeOp("item","edit/delete","batch","a batch request"));',
      '        return operations;',
      '      }',
      '      var text=asString(body);',
      '      var direct=/(?:^|\\r?\\n)(DELETE|PATCH|MERGE|PUT)\\s+([^\\s]+)\\s+HTTP\\/1\\.[01]/gi;',
      '      var match;',
      '      while((match=direct.exec(text))){ operations=operations.concat(analyzeSingle(new URL(match[2],origin).toString(),{method:match[1]})); }',
      '      var override=/(?:^|\\r?\\n)POST\\s+([^\\s]+)\\s+HTTP\\/1\\.[01][\\s\\S]*?\\r?\\nX-HTTP-Method(?:-Override)?:\\s*(DELETE|MERGE|PATCH|PUT)/gi;',
      '      while((match=override.exec(text))){ operations=operations.concat(analyzeSingle(new URL(match[1],origin).toString(),{method:match[2]})); }',
      '      return operations;',
      '    };',
      '    var prune=function(key,now){',
      '      var existing=history[key]||[];',
      '      history[key]=existing.filter(function(ts){ return now-ts<WINDOW_MS; });',
      '      return history[key];',
      '    };',
      '    var record=function(operations){',
      '      var now=Date.now();',
      '      operations.forEach(function(op){ if(op.kind!=="item") return; var list=prune(op.targetKey,now); list.push(now); });',
      '    };',
      '    var summarize=function(operations){',
      '      var full=operations.filter(function(op){ return op.kind==="full-delete"; })[0];',
      '      if(full) return "This embedded app is trying to delete the entire "+full.targetLabel+".";',
      '      var now=Date.now();',
      '      var counts={}; var labels={};',
      '      operations.forEach(function(op){ if(op.kind!=="item") return; counts[op.targetKey]=(counts[op.targetKey]||0)+1; labels[op.targetKey]=op.targetLabel; });',
      '      var reason="";',
      '      Object.keys(counts).some(function(key){ var current=prune(key,now).length; if(current+counts[key]>LIMIT){ reason="This embedded app is trying to edit or delete "+(current+counts[key])+" items from "+labels[key]+" within one minute."; return true; } return false; });',
      '      return reason;',
      '    };',
      '    var confirmRequest=function(url,init,channel){',
      '      if(destructiveCommandsApproved) return true;',
      '      var operations=analyzeSingle(url,init||{});',
      '      if(!operations.length) return true;',
      '      var reason=summarize(operations);',
      '      if(!reason){ record(operations); return true; }',
      '      var message="Firepit confirmation required\\n\\n"+reason+"\\n\\nAllow this command to be sent?";',
      '      var approved=false;',
      '      try{ approved=window.confirm(message); }catch(e){ approved=false; }',
      '      if(approved){ approveForSession(); record(operations); return true; }',
      '      try{ console.warn("Blocked by Firepit: destructive command was not confirmed.",{url:url,channel:channel,operations:operations}); }catch(e){}',
      '      return false;',
      '    };',
      '    window.__firepitDestructiveCommandGuard={ confirmRequest:confirmRequest, analyzeRequest:function(url,init){ return analyzeSingle(url,init||{}); } };',
      '  } catch(e){ console.warn("Firepit destructive-command guard init failed:",e); }',
      '})();'
    ].join('');
  }

  /**
   * Link-hint guard: neutralize <link> tags with rel values that can cause
   * outbound network requests such as dns-prefetch, preconnect, and preload.
   */
  private _buildLinkHintGuardScript(): string {
    return [
      '(function(){',
      '  try {',
      '    var allowed = new Set(["stylesheet","icon","shortcut","apple-touch-icon","apple-touch-icon-precomposed","mask-icon","canonical","license","help","author","search","alternate"]);',
      '    var normalize = function(v){ return (v || "").toLowerCase().split(/\\s+/).filter(Boolean); };',
      '    var shouldNeutralize = function(el){',
      '      if (!el || el.tagName !== "LINK") return false;',
      '      var tokens = normalize(el.getAttribute("rel"));',
      '      if (!tokens.length) return true;',
      '      return tokens.some(function(t){ return !allowed.has(t); });',
      '    };',
      '    var neutralize = function(el){',
      '      if (!shouldNeutralize(el)) return;',
      '      el.setAttribute("data-firepit-blocked-rel", el.getAttribute("rel") || "");',
      '      el.removeAttribute("href");',
      '      el.removeAttribute("rel");',
      '      if (el.parentNode) el.parentNode.removeChild(el);',
      '    };',
      '    document.querySelectorAll("link[rel]").forEach(neutralize);',
      '    new MutationObserver(function(mutations){',
      '      for (var i = 0; i < mutations.length; i++) {',
      '        var m = mutations[i];',
      '        if (m.type === "childList") {',
      '          m.addedNodes.forEach(function(node){',
      '            if (!node || node.nodeType !== 1) return;',
      '            if (node.matches && node.matches("link[rel]")) neutralize(node);',
      '            if (node.querySelectorAll) node.querySelectorAll("link[rel]").forEach(neutralize);',
      '          });',
      '        } else if (m.type === "attributes") { neutralize(m.target); }',
      '      }',
      '    }).observe(document.documentElement || document, { subtree:true, childList:true, attributes:true, attributeFilter:["rel","href"] });',
      '  } catch(e){ console.warn("Firepit link-hint guard init failed:", e); }',
      '})();'
    ].join('');
  }

  /**
   * Query-param guard: prevents data exfiltration through URL query parameters.
   * Sanitizes anchors, form actions, iframe srcs, meta refreshes, location calls,
   * history updates, and dynamically inserted DOM nodes.
   */
  private _buildQueryParamGuardScript(spOrigin: string, spSiteUrl: string, mcasWrappedSharePointOrigin: string | undefined): string {
    const safeOrigin = JSON.stringify(spOrigin);
    const safeSiteUrl = JSON.stringify(spSiteUrl);
    const safeMcasWrappedOrigin = JSON.stringify(mcasWrappedSharePointOrigin || '');
    return [
      '(function(){',
      '  try {',
      '    var MAILTO_ALLOWED = new Set(["subject","body","cc","bcc"]);',
      '    var SP_ORIGIN = ' + safeOrigin + ';',
      '    var SP_SITE_URL = ' + safeSiteUrl + ';',
      '    var MCAS_WRAPPED_SP_ORIGIN = ' + safeMcasWrappedOrigin + ';',
      '    var currentDoc = new URL(location.href); currentDoc.search = ""; currentDoc.hash = "";',
      '    var CURRENT_DOC_URL = currentDoc.toString();',
      '    var normalizePath = function(v){ var p = String(v||"/").trim(); if(!p.startsWith("/")) p="/"+p; p=p.replace(/\\/+/g,"/").replace(/\\/+$/,""); return p||"/"; };',
      '    var SP_SITE_PATH = (function(){ try { if(!SP_ORIGIN) return ""; if(!SP_SITE_URL) return "/"; return normalizePath(new URL(SP_SITE_URL,SP_ORIGIN).pathname||"/"); } catch(e){ return "/"; } })();',
      '    var isMailto = function(v){ return /^\\s*mailto:/i.test(String(v||"")); };',
      '    var resolveUrl = function(v,b){ return new URL(String(v||""),b||document.baseURI||location.href); };',
      '    var isSameDocumentTarget = function(r){ try { return r.origin===currentDoc.origin && r.pathname===currentDoc.pathname; } catch(e){ return false; } };',
      '    var isSharePointAllowedUrl = function(r){',
      '      if(!SP_ORIGIN||!r) return false;',
      '      try { var o=String(r.origin||"").toLowerCase(); if(o!==String(SP_ORIGIN).toLowerCase()) return false;',
      '        var p=normalizePath(r.pathname||"/"); if(!SP_SITE_PATH||SP_SITE_PATH==="/") return true;',
      '        return p===SP_SITE_PATH||p.startsWith(SP_SITE_PATH+"/");',
      '      } catch(e){ return false; }',
      '    };',
      '    var isTrustedStreamEmbedUrl = function(r){',
      '      if(!r) return false;',
      '      try {',
      '        var o = String(r.origin || "").toLowerCase();',
      '        var p = normalizePath(r.pathname || "/");',
      '        if(!p.endsWith("/_layouts/15/embed.aspx")) return false;',
      '        return o === String(SP_ORIGIN || "").toLowerCase() || o === String(MCAS_WRAPPED_SP_ORIGIN || "").toLowerCase();',
      '      } catch(e){ return false; }',
      '    };',
      '    var sanitizeMailto = function(v){',
      '      try { var p=new URL(String(v||"")); var k=new URLSearchParams();',
      '        for(var pair of p.searchParams.entries()){ if(MAILTO_ALLOWED.has(String(pair[0]||"").toLowerCase())) k.append(pair[0],pair[1]); }',
      '        var q=k.toString(); return "mailto:"+p.pathname+(q?"?"+q:"");',
      '      } catch(e){ return "mailto:"; }',
      '    };',
      '    var sanitizeNavigationUrl = function(v,b,mode){',
      '      var raw=String(v||""); if(!raw) return raw; if(raw.charAt(0)==="#") return raw;',
      '      if(isMailto(raw)) return sanitizeMailto(raw);',
      '      try { var parsed=resolveUrl(raw,b); parsed.username=""; parsed.password=""; parsed.search="";',
      '        if(isSharePointAllowedUrl(parsed)) return parsed.toString();',
      '        if(isSameDocumentTarget(parsed)) return parsed.toString();',
      '        if(mode==="anchor"||mode==="open") return parsed.toString();',
      '        return mode==="anchor"?"#":CURRENT_DOC_URL;',
      '      } catch(e){ return mode==="anchor"?"#":CURRENT_DOC_URL; }',
      '    };',
      '    var sanitizeAnchorHref = function(v,b){ return sanitizeNavigationUrl(v,b,"anchor"); };',
      '    var sanitizeEmbeddedUrl = function(v,b){',
      '      var raw=String(v||""); if(!raw) return "about:blank";',
      '      if(/^\\s*javascript:/i.test(raw)||/^\\s*data:/i.test(raw)) return "about:blank";',
      '      try { var p=resolveUrl(raw,b); p.username=""; p.password=""; if(isTrustedStreamEmbedUrl(p)) return p.toString(); p.search=""; if(isSameDocumentTarget(p)) return p.toString(); return "about:blank"; } catch(e){ return "about:blank"; }',
      '    };',
      '    var sanitizeSrcdocValue = function(v){',
      '      var raw=String(v||""); if(!raw) return "";',
      "      if(/http-equiv\\s*=\\s*[\"']?refresh|<base\\b|window\\s*\\.\\s*location|location\\s*\\.\\s*(href|assign|replace)|window\\s*\\.\\s*open/i.test(raw))",
      '        return "<!doctype html><meta charset=\\"utf-8\\"><title>Blocked</title>";',
      '      return raw;',
      '    };',
      '    var neutralizeBaseHref = function(el){ if(!el||el.tagName!=="BASE") return; var r=el.getAttribute("href"); if(!r) return; el.setAttribute("data-firepit-blocked-base",r); el.removeAttribute("href"); };',
      '    var isDownloadBypassHref = function(el,raw){ if(!el||el.tagName!=="A"||!el.hasAttribute("download")) return false; var h=String(raw||"").trim(); return /^blob:/i.test(h)||/^data:/i.test(h); };',
      '    var updateBlockedNavAttr = function(el,raw,b,safe){',
      '      if(!el||el.tagName!=="A") return;',
      '      if(safe==="#"){ try { var p=resolveUrl(raw,b); p.username=""; p.password=""; p.search=""; el.setAttribute("data-firepit-blocked-href",p.toString()); } catch(e){} }',
      '      else el.removeAttribute("data-firepit-blocked-href");',
      '    };',
      '    var sanitizeAttribute = function(el,attr){',
      '      if(!el||!el.getAttribute) return; var raw=el.getAttribute(attr); if(!raw) return;',
      '      if(attr==="href"&&isDownloadBypassHref(el,raw)) return;',
      '      if(attr==="href"&&el.tagName==="BASE"){ neutralizeBaseHref(el); return; }',
      '      var safe=raw, base=document.baseURI||location.href;',
      '      if(attr==="href"&&el.tagName==="A"){ safe=sanitizeAnchorHref(raw,base); updateBlockedNavAttr(el,raw,base,safe); }',
      '      if(attr==="action"&&el.tagName==="FORM") safe=sanitizeNavigationUrl(raw,base,"form");',
      '      if(attr==="src"&&/^(IFRAME|FRAME|EMBED)$/.test(el.tagName)) safe=sanitizeEmbeddedUrl(raw,base);',
      '      if(attr==="data"&&el.tagName==="OBJECT") safe=sanitizeEmbeddedUrl(raw,base);',
      '      if(attr==="srcdoc"&&/^(IFRAME|FRAME)$/.test(el.tagName)) safe=sanitizeSrcdocValue(raw);',
      '      if(attr==="http-equiv"&&el.tagName==="META"&&String(raw||"").toLowerCase()==="refresh"){',
      '        var c=String(el.getAttribute("content")||""); if(c) el.setAttribute("data-firepit-blocked-refresh",c);',
      '        el.removeAttribute("http-equiv"); el.setAttribute("content",""); return;',
      '      }',
      '      if(attr==="content"&&el.tagName==="META"&&String(el.getAttribute("http-equiv")||"").toLowerCase()==="refresh"){',
      '        el.setAttribute("data-firepit-blocked-refresh",raw); el.removeAttribute("http-equiv"); safe="";',
      '      }',
      '      if(safe!==raw) el.setAttribute(attr,safe);',
      '    };',
      '    var sanitizeFormBeforeSubmit = function(form){',
      '      if(!form||form.tagName!=="FORM") return {blocked:false};',
      '      sanitizeAttribute(form,"action");',
      '      var method=String((form.getAttribute("method")||form.method||"get")).trim().toLowerCase();',
      '      if(!method||method==="get") return {blocked:true,reason:"Blocked GET form submission by Firepit guard."};',
      '      return {blocked:false};',
      '    };',
      '    var sanitizeNode = function(node){',
      '      if(!node||node.nodeType!==1) return;',
      '      if(node.matches){',
      '        if(node.matches("a[href]")) sanitizeAttribute(node,"href");',
      '        if(node.matches("base[href]")) sanitizeAttribute(node,"href");',
      '        if(node.matches("form[action]")) sanitizeAttribute(node,"action");',
      '        if(node.matches("meta[http-equiv][content]")) sanitizeAttribute(node,"content");',
      '        if(node.matches("iframe[src],frame[src],embed[src]")) sanitizeAttribute(node,"src");',
      '        if(node.matches("object[data]")) sanitizeAttribute(node,"data");',
      '        if(node.matches("iframe[srcdoc],frame[srcdoc]")) sanitizeAttribute(node,"srcdoc");',
      '      }',
      '      if(node.querySelectorAll){',
      '        node.querySelectorAll("a[href]").forEach(function(a){sanitizeAttribute(a,"href");});',
      '        node.querySelectorAll("base[href]").forEach(function(b){sanitizeAttribute(b,"href");});',
      '        node.querySelectorAll("form[action]").forEach(function(f){sanitizeAttribute(f,"action");});',
      '        node.querySelectorAll("meta[http-equiv][content]").forEach(function(m){sanitizeAttribute(m,"content");});',
      '        node.querySelectorAll("iframe[src],frame[src],embed[src]").forEach(function(el){sanitizeAttribute(el,"src");});',
      '        node.querySelectorAll("object[data]").forEach(function(o){sanitizeAttribute(o,"data");});',
      '        node.querySelectorAll("iframe[srcdoc],frame[srcdoc]").forEach(function(el){sanitizeAttribute(el,"srcdoc");});',
      '      }',
      '    };',
      '    sanitizeNode(document.documentElement||document);',
      '    try {',
      '      var lp=window.Location&&window.Location.prototype;',
      '      if(lp&&typeof lp.assign==="function"){ var assign0=lp.assign; lp.assign=function(url){ return assign0.call(this,sanitizeNavigationUrl(url,this&&this.href?this.href:location.href,"location")); }; }',
      '      if(lp&&typeof lp.replace==="function"){ var replace0=lp.replace; lp.replace=function(url){ return replace0.call(this,sanitizeNavigationUrl(url,this&&this.href?this.href:location.href,"location")); }; }',
      '      var hrefDesc=lp?Object.getOwnPropertyDescriptor(lp,"href"):null;',
      '      if(hrefDesc&&typeof hrefDesc.set==="function"&&typeof hrefDesc.get==="function"){',
      '        Object.defineProperty(lp,"href",{configurable:true,enumerable:hrefDesc.enumerable,get:function(){return hrefDesc.get.call(this);},',
      '          set:function(v){ return hrefDesc.set.call(this,sanitizeNavigationUrl(v,this&&this.href?this.href:location.href,"location")); }});',
      '      }',
      '    } catch(e){}',
      '    if(history&&typeof history.pushState==="function"){ var push0=history.pushState.bind(history); history.pushState=function(s,t,u){ return push0(s,t,(typeof u==="string"&&u.length)?sanitizeNavigationUrl(u,location.href,"history"):u); }; }',
      '    if(history&&typeof history.replaceState==="function"){ var rpl0=history.replaceState.bind(history); history.replaceState=function(s,t,u){ return rpl0(s,t,(typeof u==="string"&&u.length)?sanitizeNavigationUrl(u,location.href,"history"):u); }; }',
      '    document.addEventListener("click",function(ev){ var t=ev.target&&ev.target.closest?ev.target.closest("a[href]"):null; if(t){ var rh=t.getAttribute("href")||""; if(!isDownloadBypassHref(t,rh)) sanitizeAttribute(t,"href"); } },true);',
      '    document.addEventListener("submit",function(ev){ var f=ev.target; if(!f||!f.matches||!f.matches("form")) return; var r=sanitizeFormBeforeSubmit(f); if(r.blocked){ ev.preventDefault(); ev.stopImmediatePropagation(); console.warn(r.reason); } },true);',
      '    var fp=window.HTMLFormElement&&window.HTMLFormElement.prototype;',
      '    if(fp&&typeof fp.submit==="function"){ var sub0=fp.submit; fp.submit=function(){ var r=sanitizeFormBeforeSubmit(this); if(r.blocked){console.warn(r.reason);return;} return sub0.call(this); }; }',
      '    if(fp&&typeof fp.requestSubmit==="function"){ var rsub0=fp.requestSubmit; fp.requestSubmit=function(sm){ var r=sanitizeFormBeforeSubmit(this); if(r.blocked){console.warn(r.reason);return;} return rsub0.call(this,sm); }; }',
      '    new MutationObserver(function(mutations){',
      '      for(var i=0;i<mutations.length;i++){',
      '        var m=mutations[i];',
      '        if(m.type==="childList"){ m.addedNodes.forEach(function(n){sanitizeNode(n);}); }',
      '        else if(m.type==="attributes"){ sanitizeNode(m.target); }',
      '      }',
      '    }).observe(document.documentElement||document,{subtree:true,childList:true,attributes:true,',
      '      attributeFilter:["href","action","method","src","data","srcdoc","content","http-equiv"]});',
      '  } catch(e){ console.warn("Firepit query-param guard init failed:",e); }',
      '})();'
    ].join('');
  }

  /**
   * Network guard: intercept fetch, XMLHttpRequest, WebSocket, EventSource and
   * navigator.sendBeacon so that only requests to the SharePoint origin are permitted.
   */
  private _buildNetworkGuardScript(spOrigin: string, mcasWrappedSharePointOrigin: string | undefined, dataverseOrigin: string | undefined): string {
    const safeOrigin = JSON.stringify(spOrigin.toLowerCase());
    const safeMcasWrappedOrigin = JSON.stringify((mcasWrappedSharePointOrigin || '').toLowerCase());
    const safeDataverseOrigin = JSON.stringify((dataverseOrigin || '').toLowerCase());
    return [
      '(function(){',
      '  try {',
      '    var SP_ORIGIN = ' + safeOrigin + ';',
      '    var MCAS_WRAPPED_SP_ORIGIN = ' + safeMcasWrappedOrigin + ';',
      '    var DATAVERSE_ORIGIN = ' + safeDataverseOrigin + ';',
      '    var MCAS_ORIGIN = "https://mcas-proxyweb.mcas-gov.us";',
      '    var MCAS_CDN_ORIGIN = "https://inline.cdn.mcas-gov.us";',
      '    var warnLast=""; var warnTimer=null;',
      '    var warn = function(kind,url,reason){',
      '      var target=url; try { var p=new URL(url,document.baseURI||location.href); if(p.origin&&p.origin!=="null") target=p.origin; } catch(e){}',
      '      var msg="Blocked outbound "+kind+(target?" to "+target:"")+". "+reason;',
      '      var key=kind+"|"+target; if(warnLast===key) return; warnLast=key;',
      '      var host=document.getElementById("firepit-network-warning");',
      '      if(!host){ host=document.createElement("div"); host.id="firepit-network-warning"; host.setAttribute("role","alert");',
      '        host.style.cssText="position:fixed;top:12px;right:12px;max-width:min(420px,calc(100vw - 24px));z-index:2147483647;pointer-events:none;font-family:system-ui,sans-serif;background:#7f1d1d;color:#fff;border:1px solid rgba(255,255,255,.18);border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.35);padding:10px 12px;font-size:13px;line-height:1.4";',
      '        (document.body||document.documentElement).appendChild(host); }',
      '      host.textContent=msg; host.hidden=false;',
      '      if(warnTimer) clearTimeout(warnTimer);',
      '      warnTimer=setTimeout(function(){ var el=document.getElementById("firepit-network-warning"); if(el) el.hidden=true; },6500);',
      '      try { console.warn(msg); } catch(e){}',
      '    };',
      '    var isAllowed = function(url){',
      '      try {',
      '        var parsed = new URL(String(url), document.baseURI || location.href);',
      '        var proto = String(parsed.protocol || "").toLowerCase();',
      '        if (proto === "data:" || proto === "blob:") return true;',
      '        if (!["http:","https:","ws:","wss:"].includes(proto)) return true;',
      '        var o = String(parsed.origin || "").toLowerCase();',
      '        return o === SP_ORIGIN || o === MCAS_WRAPPED_SP_ORIGIN || o === MCAS_ORIGIN || o === MCAS_CDN_ORIGIN || (!!DATAVERSE_ORIGIN && o === DATAVERSE_ORIGIN);',
      '      } catch(e){ return false; }',
      '    };',
      '    if(typeof window.fetch==="function"){',
      '      var fetch0=window.fetch.bind(window);',
      '      window.fetch=function(input,init){',
      '        var url=""; try { if(typeof input==="string"||input instanceof URL) url=String(input); else if(input&&typeof input.url==="string") url=input.url; } catch(e){}',
      '        if(!isAllowed(url)){ warn("fetch",url,"Origin not allowed by Firepit."); return Promise.reject(new Error("Blocked by Firepit: origin not allowed.")); }',
      '        var guardInit=init||{};',
      '        try { if(typeof Request!=="undefined"&&input instanceof Request){ guardInit={method:(init&&init.method)||input.method,headers:(init&&init.headers)||input.headers,body:init&&init.body!=null?init.body:undefined}; } } catch(e){}',
      '        var guard=window.__firepitDestructiveCommandGuard;',
      '        if(guard&&typeof guard.confirmRequest==="function"&&!guard.confirmRequest(url,guardInit,"fetch")) return Promise.reject(new Error("Blocked by Firepit: command not confirmed."));',
      '        return fetch0.apply(this,arguments);',
      '      };',
      '    }',
      '    if(window.XMLHttpRequest&&window.XMLHttpRequest.prototype){',
      '      var xhrMeta=new WeakMap(); var proto=window.XMLHttpRequest.prototype;',
      '      var open0=proto.open; var send0=proto.send; var setHeader0=proto.setRequestHeader;',
      '      proto.open=function(method,url){',
      '        xhrMeta.set(this,{method:String(method||"GET"),url:String(url==null?"":url),headers:{},blocked:!isAllowed(url)});',
      '        return open0.apply(this,arguments);',
      '      };',
      '      proto.setRequestHeader=function(name,value){ var m=xhrMeta.get(this); if(m&&name) m.headers[String(name).toLowerCase()]=String(value); return setHeader0.apply(this,arguments); };',
      '      proto.send=function(){',
      '        var m=xhrMeta.get(this);',
      '        if(m&&m.blocked){ warn("xmlhttprequest",m.url,"Origin not allowed by Firepit.");',
      '          try{if(typeof this.onerror==="function") this.onerror(new Error("Blocked by Firepit."));}catch(e){}',
      '          try{this.dispatchEvent(new Event("error"));}catch(e){} return; }',
      '        var guard=window.__firepitDestructiveCommandGuard;',
      '        if(m&&guard&&typeof guard.confirmRequest==="function"&&!guard.confirmRequest(m.url,{method:m.method,headers:m.headers,body:arguments.length?arguments[0]:undefined},"xmlhttprequest")){',
      '          try{if(typeof this.onerror==="function") this.onerror(new Error("Blocked by Firepit: command not confirmed."));}catch(e){}',
      '          try{this.dispatchEvent(new Event("error"));}catch(e){} return; }',
      '        return send0.apply(this,arguments);',
      '      };',
      '    }',
      '    if(typeof window.WebSocket==="function"){',
      '      var WS0=window.WebSocket;',
      '      window.WebSocket=function(url,protocols){',
      '        if(!isAllowed(url)){ warn("websocket",String(url),"Origin not allowed by Firepit."); throw new Error("Blocked by Firepit: origin not allowed."); }',
      '        return arguments.length>1?new WS0(url,protocols):new WS0(url);',
      '      };',
      '      window.WebSocket.prototype=WS0.prototype;',
      '      try{Object.setPrototypeOf(window.WebSocket,WS0);}catch(e){}',
      '    }',
      '    if(typeof window.EventSource==="function"){',
      '      var ES0=window.EventSource;',
      '      window.EventSource=function(url,cfg){',
      '        if(!isAllowed(url)){ warn("eventsource",String(url),"Origin not allowed by Firepit."); throw new Error("Blocked by Firepit: origin not allowed."); }',
      '        return arguments.length>1?new ES0(url,cfg):new ES0(url);',
      '      };',
      '      window.EventSource.prototype=ES0.prototype;',
      '      try{Object.setPrototypeOf(window.EventSource,ES0);}catch(e){}',
      '    }',
      '    if(navigator&&typeof navigator.sendBeacon==="function"){',
      '      var sb0=navigator.sendBeacon.bind(navigator);',
      '      navigator.sendBeacon=function(url,data){',
      '        if(!isAllowed(url)){ warn("sendbeacon",String(url),"Origin not allowed by Firepit."); return false; }',
      '        return sb0(url,data);',
      '      };',
      '    }',
      '  } catch(e){ console.warn("Firepit network guard init failed:",e); }',
      '})();'
    ].join('');
  }

  /**
   * Derive the MCAS-wrapped SharePoint origin used in some Flank Speed environments.
   * Example: https://tenant.sharepoint-mil.us -> https://tenant.sharepoint-mil.us.mcas-gov.us
   */
  private _getMcasWrappedSharePointOrigin(spOrigin: string): string | undefined {
    if (!spOrigin) {
      return undefined;
    }

    try {
      const parsed = new URL(spOrigin);
      if (!parsed.hostname) {
        return undefined;
      }

      return `${parsed.protocol}//${parsed.hostname}.mcas-gov.us`;
    } catch {
      return undefined;
    }
  }

  /**
   * Inject content immediately after the opening <head> tag in the HTML document.
   * If no <head> tag is found, prepends the content before the HTML.
   */
  private _injectIntoHead(html: string, injection: string): string {
    // Try to inject right after <head> or <head ...>
    const headMatch = html.match(/<head(\s[^>]*)?>\s*/i);
    if (headMatch && headMatch.index !== undefined) {
      const insertPos = headMatch.index + headMatch[0].length;
      return html.slice(0, insertPos) + injection + html.slice(insertPos);
    }
    // Fallback: prepend before any content
    return injection + html;
  }

  /**
   * Reuse the host page's CSP nonce for inline scripts inside the iframe document.
   */
  private _getHostCspNonce(): string | undefined {
    // In modern browsers, the nonce attribute is hidden from DOM and CSS selectors
    // so querySelector('[nonce]') often fails. We must iterate and check the IDL property.
    const scripts = document.querySelectorAll('script');
    for (let i = 0; i < scripts.length; i++) {
      if (scripts[i].nonce) {
        return scripts[i].nonce;
      }
    }
    
    const styles = document.querySelectorAll('style');
    for (let i = 0; i < styles.length; i++) {
      if (styles[i].nonce) {
        return styles[i].nonce;
      }
    }
    
    // Fallback if we still can't find it
    const nonceSource = document.querySelector<HTMLElement>('[nonce]');
    const nonce = nonceSource?.nonce || nonceSource?.getAttribute('nonce') || '';
    return nonce || undefined;
  }

  /**
   * Add the host nonce to any script or style tags that do not already have one.
   */
  private _applyNonceToScriptTags(html: string, nonce: string): string {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      
      let modified = false;
      const scripts = doc.querySelectorAll('script');
      for (let i = 0; i < scripts.length; i++) {
        if (!scripts[i].hasAttribute('nonce')) {
          scripts[i].setAttribute('nonce', nonce);
          modified = true;
        }
      }
      
      const styles = doc.querySelectorAll('style');
      for (let i = 0; i < styles.length; i++) {
        if (!styles[i].hasAttribute('nonce')) {
          styles[i].setAttribute('nonce', nonce);
          modified = true;
        }
      }
      
      if (modified) {
        // Recover doctype and return fully serialized HTML
        const doctype = doc.doctype;
        let doctypeString = '';
        if (doctype) {
          doctypeString = `<!DOCTYPE ${doctype.name}` +
            (doctype.publicId ? ` PUBLIC "${doctype.publicId}"` : '') +
            (doctype.systemId ? ` "${doctype.systemId}"` : '') + `>\n`;
        } else if (html.trim().toLowerCase().startsWith('<!doctype')) {
          doctypeString = '<!DOCTYPE html>\n';
        }
        return doctypeString + doc.documentElement.outerHTML;
      }
      return html;
    } catch (e) {
      console.warn('DOMParser failed, falling back to regex', e);
      const escapedNonce = this._escapeHtmlAttribute(nonce);
      let processed = html.replace(/<script\b(?![^>]*\bnonce\s*=)/gi, `<script nonce="${escapedNonce}"`);
      processed = processed.replace(/<style\b(?![^>]*\bnonce\s*=)/gi, `<style nonce="${escapedNonce}"`);
      return processed;
    }
  }

  /**
   * Escape attribute values injected into HTML strings.
   */
  private _escapeHtmlAttribute(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  /**
   * Get the sandbox attribute value.
   * Always uses the full-access profile so apps can use all capabilities
   * within the SharePoint origin: scripts, forms, modals, popups, and downloads.
   */
  private _getSandboxAttribute(): string {
    return 'allow-scripts allow-modals allow-popups allow-popups-to-escape-sandbox allow-forms allow-same-origin allow-downloads';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneLabel('htmlConfigurationNote', {
                  text: strings.HtmlConfigurationNote
                }),
                PropertyPaneTextField('htmlFileUrl', {
                  label: strings.HtmlFileUrlLabel,
                  description: strings.HtmlFileUrlDescription,
                  placeholder: strings.HtmlFileUrlPlaceholder
                }),
                PropertyPaneTextField('htmlCode', {
                  label: strings.HtmlCodeLabel,
                  description: strings.HtmlCodeDescription,
                  placeholder: strings.HtmlCodePlaceholder,
                  multiline: true,
                  rows: 10
                }),
                PropertyPaneToggle('fullScreen', {
                  label: strings.FullScreenLabel,
                  key: 'fullScreen',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneLabel('fullScreenHotkeyWarning', {
                  text: strings.FullScreenHotkeyWarning
                }),
                PropertyPaneTextField('iframeHeight', {
                  label: strings.IframeHeightLabel,
                  description: strings.IframeHeightDescription,
                  placeholder: '600px'
                }),
                PropertyPaneTextField('dataverseEnvironmentUrl', {
                  label: strings.DataverseEnvironmentUrlLabel,
                  description: strings.DataverseEnvironmentUrlDescription,
                  placeholder: strings.DataverseEnvironmentUrlPlaceholder
                }),
                PropertyPaneToggle('lockDown', {
                  label: strings.LockDownLabel,
                  key: 'lockDown',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneLabel('lockDownDescription', {
                  text: strings.LockDownDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
