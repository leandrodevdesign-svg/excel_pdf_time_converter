(function setupAccessPage() {
  const ACCESS_KEY = "excel_pdf_time_converter_access_granted";
  const PASSWORD_HASH = "70fa656aa0391eb9ef7bbc9c7e6771ec09e7d5d7ab1fcbbde2480d21263ee79a";

  const form = document.getElementById("accessForm");
  const input = document.getElementById("accessInput");
  const status = document.getElementById("accessStatus");

  if (!form || !input || !status) {
    return;
  }

  function readStorage(storage) {
    try {
      return storage.getItem(ACCESS_KEY) === "1";
    } catch (error) {
      return false;
    }
  }

  function writeStorage(storage) {
    try {
      storage.setItem(ACCESS_KEY, "1");
      return true;
    } catch (error) {
      return false;
    }
  }

  function hasAccess() {
    return readStorage(window.localStorage) || readStorage(window.sessionStorage);
  }

  function grantAccess() {
    return writeStorage(window.localStorage) || writeStorage(window.sessionStorage);
  }

  function getNextPath() {
    const requested = new URLSearchParams(window.location.search).get("next") || "index.html";

    if (!requested || requested.startsWith("http://") || requested.startsWith("https://") || requested.startsWith("//")) {
      return "index.html";
    }

    return requested.startsWith("./") ? requested : `./${requested.replace(/^\/+/, "")}`;
  }

  async function sha256(value) {
    const encoded = new TextEncoder().encode(value);
    const digest = await window.crypto.subtle.digest("SHA-256", encoded);
    return Array.from(new Uint8Array(digest))
      .map((byte) => byte.toString(16).padStart(2, "0"))
      .join("");
  }

  function setStatus(message) {
    status.textContent = message;
  }

  async function handleSubmit(event) {
    event.preventDefault();

    const attempt = input.value.trim();
    if (!attempt) {
      setStatus("Enter the access code.");
      return;
    }

    form.classList.add("is-busy");
    setStatus("Checking...");

    try {
      const attemptHash = await sha256(attempt);
      if (attemptHash !== PASSWORD_HASH) {
        input.value = "";
        setStatus("Incorrect code.");
        input.focus();
        return;
      }

      grantAccess();
      window.location.replace(getNextPath());
    } catch (error) {
      setStatus("This browser could not validate the code.");
    } finally {
      form.classList.remove("is-busy");
    }
  }

  if (hasAccess()) {
    window.location.replace(getNextPath());
    return;
  }

  form.addEventListener("submit", handleSubmit);
  input.focus();
})();
