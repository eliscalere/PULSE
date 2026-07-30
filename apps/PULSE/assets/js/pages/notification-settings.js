PAGE_RENDERERS["notification-settings"] = function () {
  const db = window.AEWTTR.db;
  // Live state, not a draft — every toggle/tone change saves itself
  // (debounced) instead of waiting for an explicit Save click. Cancel/Save
  // buttons are gone; "current" always mirrors what's on screen and what's
  // persisted (or about to be, within the debounce window).
  const current = normalizeNotificationPrefs(db.user.notificationPrefs);

  const AREA_LABELS = NOTIFICATION_AREA_META;
  let saveTimer = null;

  setTopbar("Notification Settings", "Channels, areas, Document Review delivery, and tone. Changes save automatically.", "");

  function setSaveStatus(text, isError) {
    const el = $("#notif-save-status");
    if (!el) return;
    el.textContent = text;
    el.style.color = isError ? "var(--aewttr-red)" : "var(--aewttr-muted)";
    // data-state is watched by isBackgroundDataSwapBlocked() so a poll
    // can't replace window.AEWTTR.db while this autosave is mid-flight.
    if (isError) el.dataset.state = "error";
    else if (text === "Saving…") el.dataset.state = "saving";
    else if (text.indexOf("Saving") === 0) el.dataset.state = "pending";
    else el.dataset.state = "saved";
  }

  function prefsPayload() {
    const areas = current.everything ? NOTIFICATION_AREAS.slice() : current.areas.slice();
    return {
      areas,
      tone: current.tone,
      channels: { ...current.channels },
      everything: !!current.everything,
      documents: {
        delivery: current.documents.delivery,
        digestFrequency: current.documents.digestFrequency
      },
      digest: { ...current.digest },
      actionItemDigest: { ...current.actionItemDigest },
      defaultPage: current.defaultPage || ""
    };
  }

  async function flushPrefsNow() {
    unregisterAutosaveFlusher(flushPrefsNow);
    if (!saveTimer) return;
    clearTimeout(saveTimer);
    saveTimer = null;
    setSaveStatus("Saving…");
    const payload = prefsPayload();
    const liveDb = window.AEWTTR.db;
    liveDb.user.notificationPrefs = payload;
    const member = typeof currentUserMember === "function" ? currentUserMember() : null;
    if (member) member.notificationPrefs = payload;
    if (isSharePointMode()) {
      try {
        await sharePointAdapter.saveCurrentUserNotificationPrefs(window.AEWTTR.siteUrl, window.AEWTTR.currentSpUser, payload);
      } catch (e) {
        setSaveStatus("Couldn't save — check your connection", true);
        toast((e && e.friendly) || "Could not save notification preferences.", "error");
        return;
      }
    }
    setSaveStatus("All changes saved");
  }

  function persistPrefs() {
    setSaveStatus("Saving soon…");
    if (saveTimer) clearTimeout(saveTimer);
    registerAutosaveFlusher(flushPrefsNow);
    saveTimer = setTimeout(flushPrefsNow, 400);
  }

  function documentsAreaOn() {
    return current.everything || current.areas.includes("Documents");
  }

  function defaultPageOptionsForCurrentUser() {
    const canAdmin = typeof canCurrentUserAccessAdmin === "function" ? canCurrentUserAccessAdmin() : false;
    const canFinance = typeof canAssignTravelCo === "function" ? canAssignTravelCo() : false;
    return DEFAULT_PAGE_OPTIONS.filter((opt) => {
      if (opt.route === "admin") return canAdmin;
      if (opt.route === "travel/finance") return canFinance;
      return true;
    });
  }

  function toneCardHtml(t) {
    const active = current.tone === t.key;
    // Funny doesn't have one fixed sample — it draws from a pool of
    // openings/sign-offs at send time, so the preview re-rolls too instead
    // of always showing the same joke.
    const sample = t.key === "funny" && typeof randomFunnySampleText === "function" ? randomFunnySampleText() : t.sample;
    return `
      <button type="button" class="notif-tone-card ${active ? "active" : ""}" data-tone="${t.key}">
        <span class="notif-tone-card-head">
          <strong>${escapeHtml(t.label)}</strong>
          ${active ? `<i class="bx bx-check-circle notif-tone-check"></i>` : ""}
        </span>
        <span class="notif-tone-blurb">${escapeHtml(t.blurb)}</span>
        <span class="notif-tone-preview">${escapeHtml(sample)}</span>
      </button>
    `;
  }

  function render() {
    const docsOn = documentsAreaOn();
    const digestOn = current.documents.delivery === "digest";

    $("#page-content").innerHTML = `
      <div class="notif-settings-page">
        <div class="notif-settings-toolbar">
          <p class="notif-settings-lede">Delivery is personal — site admins can still mute Teams or email for everyone under Admin → Notifications.</p>
          <div class="notif-save-status" id="notif-save-status" data-state="saved">All changes saved</div>
        </div>

        <div class="notif-settings-grid">
          <section class="notif-panel notif-panel--channels">
            <header class="notif-panel-head">
              <h3 class="notif-panel-title">Channels</h3>
              <p class="notif-panel-sub">Independent of area filters below.</p>
            </header>
            <div class="notif-channel-row">
              <label class="notif-toggle-row notif-toggle-row--compact">
                <span class="notif-toggle-label"><i class="bx bx-envelope"></i> Email</span>
                <span class="aewttr-switch"><input type="checkbox" id="notif-ch-email" ${current.channels.email ? "checked" : ""}><span class="aewttr-switch-track"></span></span>
              </label>
              <label class="notif-toggle-row notif-toggle-row--compact">
                <span class="notif-toggle-label"><i class="bx bxl-microsoft-teams"></i> Teams</span>
                <span class="aewttr-switch"><input type="checkbox" id="notif-ch-teams" ${current.channels.teams ? "checked" : ""}><span class="aewttr-switch-track"></span></span>
              </label>
            </div>
          </section>

          <section class="notif-panel notif-panel--landing">
            <header class="notif-panel-head">
              <h3 class="notif-panel-title">Landing Page</h3>
              <p class="notif-panel-sub">Where PULSE opens when you log in.</p>
            </header>
            <label class="notif-field">
              <span class="notif-field-label">Default page</span>
              <select id="notif-default-page" class="select-aewttr">
                ${defaultPageOptionsForCurrentUser().map((opt) => `<option value="${escapeHtml(opt.route)}" ${current.defaultPage === opt.route ? "selected" : ""}>${escapeHtml(opt.label)}</option>`).join("")}
              </select>
            </label>
            <p class="notif-settings-note notif-settings-note--inline">"Automatic" uses PULSE's normal role-based default (Team Overview for Admins, Awaiting Finance for Finance Admins, Document Review for Document Admins, Dashboard otherwise).</p>
          </section>

          <section class="notif-panel notif-panel--docreview">
            <header class="notif-panel-head">
              <h3 class="notif-panel-title">Document Review</h3>
              <p class="notif-panel-sub">How review and signing events reach you when Documents is on.</p>
            </header>
            <div class="notif-delivery-mode" role="radiogroup" aria-label="Document Review delivery">
              <label class="notif-choice ${current.documents.delivery === "immediate" ? "is-active" : ""} ${!docsOn ? "is-disabled" : ""}">
                <input type="radio" name="notif-doc-delivery" value="immediate" ${current.documents.delivery === "immediate" ? "checked" : ""} ${!docsOn ? "disabled" : ""}>
                <span class="notif-choice-body">
                  <strong>Every event</strong>
                  <span>Notify on each review action as it happens.</span>
                </span>
              </label>
              <label class="notif-choice ${digestOn ? "is-active" : ""} ${!docsOn ? "is-disabled" : ""}">
                <input type="radio" name="notif-doc-delivery" value="digest" ${digestOn ? "checked" : ""} ${!docsOn ? "disabled" : ""}>
                <span class="notif-choice-body">
                  <strong>Digest</strong>
                  <span>Hold immediate Document Review pings; prefer a batch later.</span>
                </span>
              </label>
            </div>
            <div class="notif-digest-row ${digestOn && docsOn ? "" : "is-hidden"}">
              <label class="notif-field">
                <span class="notif-field-label">Frequency</span>
                <select id="notif-doc-digest-freq" class="select-aewttr" ${!digestOn || !docsOn ? "disabled" : ""}>
                  <option value="daily" ${current.documents.digestFrequency === "daily" ? "selected" : ""}>Daily</option>
                  <option value="weekly" ${current.documents.digestFrequency === "weekly" ? "selected" : ""}>Weekly</option>
                </select>
              </label>
              <p class="notif-settings-note notif-settings-note--inline">
                Frequency is saved for a future scheduled digest job. Until then, Digest mode skips immediate Document Review notifications (no batch email/Teams yet).
              </p>
            </div>
            ${!docsOn ? `<p class="notif-settings-note">Document Review area is off — enable it below (or switch to Everything) to receive these.</p>` : ""}
          </section>

          <section class="notif-panel notif-panel--areas">
            <header class="notif-panel-head notif-panel-head--row">
              <div>
                <h3 class="notif-panel-title">Apps &amp; Areas</h3>
                <p class="notif-panel-sub">Which parts of PULSE may notify you.</p>
              </div>
              <label class="notif-toggle-row notif-toggle-row--master">
                <span class="notif-toggle-label">Everything</span>
                <span class="aewttr-switch"><input type="checkbox" id="notif-everything" ${current.everything ? "checked" : ""}><span class="aewttr-switch-track"></span></span>
              </label>
            </header>
            <div class="notif-area-grid ${current.everything ? "is-dimmed" : ""}" id="notif-area-grid">
              ${NOTIFICATION_AREAS.map((area) => {
                const checked = current.everything || current.areas.includes(area);
                return `
                <label class="notif-toggle-row notif-toggle-row--compact">
                  <span class="notif-toggle-label">${escapeHtml(AREA_LABELS[area].label)}</span>
                  <span class="aewttr-switch"><input type="checkbox" data-pref-area="${escapeHtml(area)}" ${checked ? "checked" : ""} ${current.everything ? "disabled" : ""}><span class="aewttr-switch-track"></span></span>
                </label>`;
              }).join("")}
            </div>
            ${current.everything ? `<p class="notif-settings-note">Everything is on — all areas notify you. Turn it off to pick specific ones.</p>` : ""}
          </section>

          <section class="notif-panel notif-panel--tone">
            <header class="notif-panel-head">
              <h3 class="notif-panel-title">Tone</h3>
              <p class="notif-panel-sub">Greeting and sign-off only — facts stay the same.</p>
            </header>
            <div class="notif-tone-grid" id="notif-tone-grid">
              ${NOTIFICATION_TONES.map(toneCardHtml).join("")}
            </div>
          </section>

          <section class="notif-panel notif-panel--digest">
            <header class="notif-panel-head notif-panel-head--row">
              <div>
                <h3 class="notif-panel-title">Daily Digest</h3>
                <p class="notif-panel-sub">A summary of your open action items shown when you open PULSE.</p>
              </div>
              <label class="notif-toggle-row notif-toggle-row--master">
                <span class="notif-toggle-label">Enabled</span>
                <span class="aewttr-switch"><input type="checkbox" id="notif-digest-enabled" ${current.digest.enabled ? "checked" : ""}><span class="aewttr-switch-track"></span></span>
              </label>
            </header>
            <div class="notif-digest-row ${current.digest.enabled ? "" : "is-hidden"}">
              <label class="notif-field">
                <span class="notif-field-label">Frequency</span>
                <select id="notif-digest-freq" class="select-aewttr" ${!current.digest.enabled ? "disabled" : ""}>
                  <option value="daily" ${current.digest.frequency === "daily" ? "selected" : ""}>Daily (show once per day)</option>
                  <option value="weekly" ${current.digest.frequency === "weekly" ? "selected" : ""}>Weekly (show once per week)</option>
                </select>
              </label>
              <button class="btn-aewttr-outline" id="notif-digest-preview" style="margin-top:4px;"><i class="bx bx-list-check"></i> Preview Digest Now</button>
            </div>
            <div class="notif-digest-row">
              <label class="notif-toggle-row">
                <span class="notif-toggle-label">Email &amp; Teams delivery</span>
                <span class="aewttr-switch"><input type="checkbox" id="notif-actionitem-digest-enabled" ${current.actionItemDigest.enabled ? "checked" : ""}><span class="aewttr-switch-track"></span></span>
              </label>
              <p class="notif-settings-note notif-settings-note--inline">
                When on (default), you'll get a daily email and Teams message listing your own open action items — sent automatically once a day, independent of the in-app popup above.
              </p>
            </div>
          </section>

          <section class="notif-panel notif-panel--test">
            <header class="notif-panel-head">
              <h3 class="notif-panel-title">Test</h3>
              <p class="notif-panel-sub">Sample with the tone above (bypasses area/digest gates).</p>
            </header>
            <div class="notif-test-row">
              <button class="btn-aewttr-outline" id="notif-send-test"${tip("Send a sample notification to your own email/Teams with the currently selected tone")}><i class="bx bx-paper-plane"></i> Send Test Notification</button>
              <div id="notif-test-result" class="notif-settings-note notif-settings-note--inline"></div>
            </div>
          </section>
        </div>
      </div>
    `;

    $("#notif-ch-email").addEventListener("change", (e) => { current.channels.email = e.target.checked; persistPrefs(); });
    $("#notif-ch-teams").addEventListener("change", (e) => { current.channels.teams = e.target.checked; persistPrefs(); });

    $("#notif-default-page").addEventListener("change", (e) => {
      current.defaultPage = e.target.value || "";
      persistPrefs();
    });

    $("#notif-everything").addEventListener("change", (e) => {
      current.everything = e.target.checked;
      if (current.everything) current.areas = NOTIFICATION_AREAS.slice();
      render();
      persistPrefs();
    });

    $all("[data-pref-area]", $("#page-content")).forEach((cb) => {
      cb.addEventListener("change", () => {
        current.areas = $all("[data-pref-area]", $("#page-content")).filter((c) => c.checked).map((c) => c.dataset.prefArea);
        current.everything = current.areas.length === NOTIFICATION_AREAS.length;
        render();
        persistPrefs();
      });
    });

    $all('input[name="notif-doc-delivery"]', $("#page-content")).forEach((radio) => {
      radio.addEventListener("change", () => {
        if (!radio.checked) return;
        current.documents.delivery = radio.value === "digest" ? "digest" : "immediate";
        render();
        persistPrefs();
      });
    });

    const freqEl = $("#notif-doc-digest-freq");
    if (freqEl) {
      freqEl.addEventListener("change", () => {
        current.documents.digestFrequency = freqEl.value === "weekly" ? "weekly" : "daily";
        persistPrefs();
      });
    }

    $all("[data-tone]", $("#notif-tone-grid")).forEach((card) => {
      card.addEventListener("click", () => {
        current.tone = card.dataset.tone;
        render();
        persistPrefs();
      });
    });

    $("#notif-digest-enabled").addEventListener("change", (e) => {
      current.digest.enabled = e.target.checked;
      render();
      persistPrefs();
    });

    const digestFreqEl = $("#notif-digest-freq");
    if (digestFreqEl) {
      digestFreqEl.addEventListener("change", () => {
        current.digest.frequency = digestFreqEl.value === "weekly" ? "weekly" : "daily";
        persistPrefs();
      });
    }

    const digestPreviewBtn = $("#notif-digest-preview");
    if (digestPreviewBtn) {
      digestPreviewBtn.addEventListener("click", () => {
        if (typeof showDigestModal === "function") showDigestModal({ force: true });
        else toast("Digest not available — refresh the app.", "error");
      });
    }

    $("#notif-actionitem-digest-enabled").addEventListener("change", (e) => {
      current.actionItemDigest.enabled = e.target.checked;
      persistPrefs();
    });

    $("#notif-send-test").addEventListener("click", async () => {
      const btn = $("#notif-send-test");
      const resultEl = $("#notif-test-result");
      btn.disabled = true;
      btn.innerHTML = `<i class="bx bx-loader-alt bx-spin"></i> Sending…`;
      try {
        if (typeof sendTestNotificationForCurrentUser !== "function") throw new Error("Notification helper not loaded — refresh the app.");
        const result = await sendTestNotificationForCurrentUser(current.tone);
        const parts = [];
        parts.push(result.email && result.email.sent ? "email sent" : `email not sent${result.email && result.email.error ? ` (${result.email.error})` : ""}`);
        parts.push(result.teams && result.teams.queued ? "Teams queued" : `Teams not queued${result.teams && result.teams.error ? ` (${result.teams.error})` : ""}`);
        if (resultEl) resultEl.textContent = `Test complete: ${parts.join(" · ")}. Check your inbox and Teams.`;
        toast("Test notification sent", "success");
      } catch (e) {
        const message = (e && e.message) || String(e);
        if (resultEl) resultEl.textContent = message;
        toast(message, "error");
      }
      btn.disabled = false;
      btn.innerHTML = `<i class="bx bx-paper-plane"></i> Send Test Notification`;
    });
  }

  render();
};
