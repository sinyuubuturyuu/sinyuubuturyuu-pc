(function () {
  const FULL_WINDOW_QUERY_KEY = "windowMode";
  const FULL_WINDOW_QUERY_VALUE = "full";
  const WINDOW_TOKEN_QUERY_KEY = "windowToken";
  const RETURN_CHANNEL_NAME = "launcher-return-channel";
  const openedInspectionWindows = new Map();
  const returnChannel = typeof BroadcastChannel === "function"
    ? new BroadcastChannel(RETURN_CHANNEL_NAME)
    : null;

  function appendFullWindowParam(rawUrl) {
    const targetUrl = new URL(rawUrl, window.location.href);
    targetUrl.searchParams.set(FULL_WINDOW_QUERY_KEY, FULL_WINDOW_QUERY_VALUE);
    return targetUrl.toString();
  }

  function appendWindowTokenParam(rawUrl, token) {
    const targetUrl = new URL(rawUrl, window.location.href);
    targetUrl.searchParams.set(WINDOW_TOKEN_QUERY_KEY, token);
    return targetUrl.toString();
  }

  function createWindowToken() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }

    return `inspection-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function buildPopupFeatures() {
    const width = Math.max(window.screen?.availWidth || 1280, 960);
    const height = Math.max(window.screen?.availHeight || 720, 720);
    const left = Math.max(Math.floor(((window.screen?.availWidth || width) - width) / 2), 0);
    const top = Math.max(Math.floor(((window.screen?.availHeight || height) - height) / 2), 0);

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

  function openInspectionWindow(rawUrl) {
    const windowToken = createWindowToken();
    const targetUrl = appendWindowTokenParam(appendFullWindowParam(rawUrl), windowToken);
    const popup = window.open("", "_blank", buildPopupFeatures());

    if (popup && !popup.closed) {
      try {
        openedInspectionWindows.set(windowToken, popup);
        popup.location.replace(targetUrl);
        popup.focus();
        return true;
      } catch (error) {
        console.warn("Failed to open popup window:", error);
      }
    }

    window.location.href = targetUrl;
    return false;
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

    return targetWindow.closed;
  }

  function cleanupClosedInspectionWindows() {
    for (const [token, popup] of openedInspectionWindows.entries()) {
      if (!popup || popup.closed) {
        openedInspectionWindows.delete(token);
      }
    }
  }

  function handleReturnRequest(message) {
    if (!message || message.type !== "return-to-launcher" || !message.windowToken) {
      return;
    }

    const popup = openedInspectionWindows.get(message.windowToken);
    if (!popup) {
      return;
    }

    if (message.targetUrl) {
      try {
        window.location.href = message.targetUrl;
        window.focus();
      } catch (error) {
        console.warn("Failed to navigate launcher window:", error);
      }
    }

    closeInspectionWindow(message.windowToken);
    cleanupClosedInspectionWindows();
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

  function maximizeCurrentWindowIfRequested() {
    const params = new URLSearchParams(window.location.search);
    const isFullWindowMode = params.get(FULL_WINDOW_QUERY_KEY) === FULL_WINDOW_QUERY_VALUE;

    if (!isFullWindowMode) {
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

  function bindReturnToLauncher(link) {
    if (!link) {
      return;
    }

    link.addEventListener("click", function (event) {
      const params = new URLSearchParams(window.location.search);
      const isFullWindowMode = params.get(FULL_WINDOW_QUERY_KEY) === FULL_WINDOW_QUERY_VALUE;
      const windowToken = params.get(WINDOW_TOKEN_QUERY_KEY);
      const openerWindow = window.opener;

      if (!isFullWindowMode) {
        return;
      }

      event.preventDefault();

      const targetUrl = new URL(link.href, window.location.href).toString();

      if (returnChannel && windowToken) {
        returnChannel.postMessage({
          type: "return-to-launcher",
          targetUrl: targetUrl,
          windowToken: windowToken
        });
      }

      if (!window.closed) {
        closeInspectionWindow(window);
      }

      if (!window.closed) {
        try {
          window.open("", "_self");
        } catch (error) {
          console.warn("Failed to reopen current window before close:", error);
        }
        closeInspectionWindow(window);
      }

      try {
        if (openerWindow && !openerWindow.closed) {
          if (openerWindow.launcherWindow
            && typeof openerWindow.launcherWindow.closeInspectionWindow === "function") {
            openerWindow.launcherWindow.closeInspectionWindow(window);
          }
          openerWindow.location.href = targetUrl;
          openerWindow.focus();
        }
      } catch (error) {
        console.warn("Failed to focus launcher window:", error);
      }

      if (!window.closed) {
        window.location.href = targetUrl;
      }
    });
  }

  window.addEventListener("beforeunload", cleanupClosedInspectionWindows);

  if (returnChannel) {
    returnChannel.addEventListener("message", function (event) {
      handleReturnRequest(event.data);
    });
  }

  window.launcherWindow = {
    bindInspectionLaunch: bindInspectionLaunch,
    maximizeCurrentWindowIfRequested: maximizeCurrentWindowIfRequested,
    bindReturnToLauncher: bindReturnToLauncher,
    closeInspectionWindow: closeInspectionWindow
  };
})();
