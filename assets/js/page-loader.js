(function registerPageLoader() {
  function hideInitialLoader(delay) {
    const pageLoader = document.getElementById("pageLoader");

    if ("scrollRestoration" in window.history) {
      window.history.scrollRestoration = "manual";
    }

    window.scrollTo({ top: 0, left: 0, behavior: "auto" });

    window.setTimeout(function finalizeLoader() {
      if (pageLoader) {
        pageLoader.classList.add("is-hidden");
      }

      document.body.classList.remove("is-loading");
      window.scrollTo({ top: 0, left: 0, behavior: "auto" });
    }, typeof delay === "number" ? delay : 550);
  }

  window.addEventListener("load", function onWindowLoad() {
    hideInitialLoader(550);
  }, { once: true });
})();
