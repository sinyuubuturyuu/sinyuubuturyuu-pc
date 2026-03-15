(function () {
  const FULL_WINDOW_QUERY_KEY = "windowMode";
  const FULL_WINDOW_QUERY_VALUE = "full";

  function appendFullWindowParam(rawUrl) {
    const targetUrl = new URL(rawUrl, window.location.href);
    targetUrl.searchParams.set(FULL_WINDOW_QUERY_KEY, FULL_WINDOW_QUERY_VALUE);
    return targetUrl.toString();
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
    const targetUrl = appendFullWindowParam(rawUrl);
    const popup = window.open("", "_blank", buildPopupFeatures());

    if (popup && !popup.closed) {
      try {
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

  window.launcherWindow = {
    bindInspectionLaunch: bindInspectionLaunch,
    maximizeCurrentWindowIfRequested: maximizeCurrentWindowIfRequested
  };
})();
