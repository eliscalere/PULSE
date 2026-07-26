import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { IAewttrWorkspaceWebPartProps } from './IAewttrWorkspaceWebPartProps';
import { AewttrWorkspaceApp } from './components/AewttrWorkspaceApp';

export default class AewttrWorkspaceWebPart extends BaseClientSideWebPart<IAewttrWorkspaceWebPartProps> {
  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement = React.createElement(AewttrWorkspaceApp, {
      context: this.context,
      isDarkTheme: this._isDarkTheme,
      description: this.properties.description
    });

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    this._isDarkTheme = Boolean(currentTheme?.isInverted);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
