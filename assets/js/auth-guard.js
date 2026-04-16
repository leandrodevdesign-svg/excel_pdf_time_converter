(function guardProtectedPages() {
  const ACCESS_PAGE = "access.html";
  const ACCESS_KEY = "excel_pdf_time_converter_access_granted";

  function readStorage(storage) {
    try {
      return storage.getItem(ACCESS_KEY) === "1";
    } catch (error) {
      return false;
    }
  }

  function hasAccess() {
    return readStorage(window.localStorage) || readStorage(window.sessionStorage);
  }

  const currentFile = window.location.pathname.split("/").pop() || "index.html";

  if (currentFile === ACCESS_PAGE) {
    return;
  }

  if (hasAccess()) {
    return;
  }

  const nextPath = `${currentFile}${window.location.search}${window.location.hash}`;
  const targetUrl = new URL(`./${ACCESS_PAGE}`, window.location.href);
  targetUrl.searchParams.set("next", nextPath || "index.html");
  window.location.replace(targetUrl.toString());
})();
