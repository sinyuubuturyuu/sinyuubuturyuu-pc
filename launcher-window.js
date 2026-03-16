(function () {
  const FULL_WINDOW_QUERY_KEY = "windowMode";
  const FULL_WINDOW_QUERY_VALUE = "full";
  const LAUNCHER_WINDOW_QUERY_VALUE = "launcher";
  const WINDOW_TOKEN_QUERY_KEY = "windowToken";
  const RETURN_CHANNEL_NAME = "launcher-return-channel";
  const RETURN_MESSAGE_TYPE = "return-to-launcher";

  const openedInspectionWindows = new Map();
  const returnChannel = typeof BroadcastChannel === "function"
    ? new BroadcastChannel(RETURN_CHANNEL_NAME)
    : null;

  function createWindowToken() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }

    return `inspection-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function resolveUrl(rawUrl) {
    return new URL(rawUrl, window.location.href);
  }

  function createInspectionWindowUrl(rawUrl, windowToken, windowMode = FULL_WINDOW_QUERY_VALUE) {
    const targetUrl = resolveUrl(rawUrl);
    targetUrl.searchParams.set(FULL_WINDOW_QUERY_KEY, windowMode);
    targetUrl.searchParams.set(WINDOW_TOKEN_QUERY_KEY, windowToken);
    return targetUrl.toString();
  }

  function getWindowParams(search = window.location.search) {
    return new URLSearchParams(search);
  }

  function isFullWindowMode(search = window.location.search) {
    return getWindowParams(search).get(FULL_WINDOW_QUERY_KEY) === FULL_WINDOW_QUERY_VALUE;
  }

  function isManagedChildWindow(search = window.location.search) {
    return Boolean(getWindowToken(search));
  }

  function getWindowToken(search = window.location.search) {
    return getWindowParams(search).get(WINDOW_TOKEN_QUERY_KEY) || "";
  }

  function buildPopupFeatures(windowMode = FULL_WINDOW_QUERY_VALUE) {
    const isLauncherSizeMode = windowMode === LAUNCHER_WINDOW_QUERY_VALUE;
    const width = isLauncherSizeMode
      ? Math.max(window.outerWidth || window.innerWidth || 960, 640)
      : Math.max(window.screen?.availWidth || 1280, 960);
    const height = isLauncherSizeMode
      ? Math.max(window.outerHeight || window.innerHeight || 720, 480)
      : Math.max(window.screen?.availHeight || 720, 720);
    const left = isLauncherSizeMode
      ? Math.max(window.screenX || window.screenLeft || 0, 0)
      : Math.max(Math.floor(((window.screen?.availWidth || width) - width) / 2), 0);
    const top = isLauncherSizeMode
      ? Math.max(window.screenY || window.screenTop || 0, 0)
      : Math.max(Math.floor(((window.screen?.availHeight || height) - height) / 2), 0);

    return [
      "popup=yes",
      "resizable=yes",
      "scrollbars=yes",
      `width=${width}`,
      `height=${height}`,
      `left=${left}`,
      `top=${top}`
    ].join(",");
  }

  function cleanupClosedInspectionWindows() {
    for (const [token, popup] of openedInspectionWindows.entries()) {
      if (!popup || popup.closed) {
        openedInspectionWindows.delete(token);
      }
    }
  }

  function rememberInspectionWindow(windowToken, popup) {
    cleanupClosedInspectionWindows();
    openedInspectionWindows.set(windowToken, popup);
  }

  function closeInspectionWindow(targetWindowOrToken) {
    const targetWindow = typeof targetWindowOrToken === "string"
      ? openedInspectionWindows.get(targetWindowOrToken)
      : targetWindowOrToken;

    if (!targetWindow || targetWindow.closed) {
      return false;
    }

    try {
      targetWindow.close();
    } catch (error) {
      console.warn("Failed to close inspection window:", error);
    }

    cleanupClosedInspectionWindows();
    return targetWindow.closed;
  }

  function navigateCurrentWindow(targetUrl) {
    window.location.href = targetUrl;
  }

  function focusLauncherWindow(targetUrl) {
    const openerWindow = window.opener;

    if (!openerWindow || openerWindow.closed) {
      return false;
    }

    try {
      openerWindow.location.href = targetUrl;
      openerWindow.focus();
      return true;
    } catch (error) {
      console.warn("Failed to focus launcher window:", error);
      return false;
    }
  }

  function createReturnMessage(targetUrl, windowToken) {
    return {
      type: RETURN_MESSAGE_TYPE,
      targetUrl: targetUrl,
      windowToken: windowToken
    };
  }

  function handleReturnRequest(message) {
    if (!message || message.type !== RETURN_MESSAGE_TYPE || !message.windowToken) {
      return false;
    }

    if (message.targetUrl) {
      navigateCurrentWindow(message.targetUrl);
      window.focus();
    }

    closeInspectionWindow(message.windowToken);
    return true;
  }

  function notifyLauncherReturn(targetUrl, windowToken) {
    const message = createReturnMessage(targetUrl, windowToken);
    let handled = false;

    try {
      const openerWindow = window.opener;
      if (openerWindow
        && !openerWindow.closed
        && openerWindow.launcherWindow
        && typeof openerWindow.launcherWindow.handleReturnRequest === "function") {
        handled = openerWindow.launcherWindow.handleReturnRequest(message) || handled;
      }
    } catch (error) {
      console.warn("Failed to call launcher return handler:", error);
    }

    if (returnChannel && windowToken) {
      returnChannel.postMessage(message);
      handled = true;
    }

    return handled;
  }

  function closeCurrentInspectionWindow() {
    if (closeInspectionWindow(window)) {
      return true;
    }

    try {
      window.open("", "_self");
    } catch (error) {
      console.warn("Failed to prepare current window for close:", error);
    }

    return closeInspectionWindow(window);
  }

  function openInspectionWindow(rawUrl, windowMode = FULL_WINDOW_QUERY_VALUE) {
    const windowToken = createWindowToken();
    const targetUrl = createInspectionWindowUrl(rawUrl, windowToken, windowMode);
    const popup = window.open("", "_blank", buildPopupFeatures(windowMode));

    if (popup && !popup.closed) {
      try {
        rememberInspectionWindow(windowToken, popup);
        popup.location.replace(targetUrl);
        popup.focus();
        return true;
      } catch (error) {
        console.warn("Failed to open popup window:", error);
      }
    }

    navigateCurrentWindow(targetUrl);
    return false;
  }

  function bindInspectionLaunch(link) {
    if (!link) {
      return;
    }

    link.addEventListener("click", function (event) {
      event.preventDefault();
      openInspectionWindow(link.href);
    });
  }

  function bindLauncherSizedLaunch(link) {
    if (!link) {
      return;
    }

    link.addEventListener("click", function (event) {
      event.preventDefault();
      openInspectionWindow(link.href, LAUNCHER_WINDOW_QUERY_VALUE);
    });
  }

  function maximizeCurrentWindowIfRequested() {
    if (!isFullWindowMode()) {
      return false;
    }

    document.documentElement.classList.add("window-mode-full");
    document.body.classList.add("window-mode-full");

    try {
      if (typeof window.moveTo === "function") {
        window.moveTo(0, 0);
      }
      if (typeof window.resizeTo === "function") {
        window.resizeTo(window.screen?.availWidth || window.outerWidth, window.screen?.availHeight || window.outerHeight);
      }
      window.focus();
    } catch (error) {
      console.warn("Failed to maximize window:", error);
    }

    return true;
  }

  // Returning from a child window needs three fallbacks:
  // 1. notify the launcher, 2. close this popup, 3. navigate if close is blocked.
  function bindReturnToLauncher(link) {
    if (!link) {
      return;
    }

    link.addEventListener("click", function (event) {
      if (!isManagedChildWindow()) {
        return;
      }

      event.preventDefault();

      const targetUrl = resolveUrl(link.href).toString();
      const windowToken = getWindowToken();

      notifyLauncherReturn(targetUrl, windowToken);
      closeCurrentInspectionWindow();

      if (!window.closed) {
        focusLauncherWindow(targetUrl);
      }

      if (!window.closed) {
        navigateCurrentWindow(targetUrl);
      }
    });
  }

  window.addEventListener("beforeunload", cleanupClosedInspectionWindows);
  window.addEventListener("pagehide", cleanupClosedInspectionWindows);
  window.addEventListener("focus", cleanupClosedInspectionWindows);

  if (returnChannel) {
    returnChannel.addEventListener("message", function (event) {
      handleReturnRequest(event.data);
    });
  }

  window.launcherWindow = {
    bindInspectionLaunch: bindInspectionLaunch,
    bindLauncherSizedLaunch: bindLauncherSizedLaunch,
    maximizeCurrentWindowIfRequested: maximizeCurrentWindowIfRequested,
    bindReturnToLauncher: bindReturnToLauncher,
    closeInspectionWindow: closeInspectionWindow,
    handleReturnRequest: handleReturnRequest
  };
})();
