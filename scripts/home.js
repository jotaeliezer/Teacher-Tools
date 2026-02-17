(() => {
  const homeShell = document.getElementById("homeShell");
  const appRoot = document.getElementById("appRoot");
  const accessForm = document.getElementById("accessForm");
  const emailInput = document.getElementById("homeEmail");
  const statusEl = document.getElementById("launchStatus");
  const messageEl = document.getElementById("emailMessage");
  const launchBtn = document.getElementById("launchBtn");
  const barFill = document.querySelector(".launch-bar-fill");
  const openers = document.querySelectorAll(".open-launcher");
  const backdrop = document.getElementById("homeBackdrop");
  const closeBtn = document.getElementById("homeCloseBtn");
  let isLaunching = false;

  if (!homeShell || !appRoot || !accessForm) return;

  let savedEmail = localStorage.getItem("teacherToolsEmail");
  if (savedEmail && emailInput) {
    emailInput.value = savedEmail;
    if (statusEl) statusEl.textContent = "Email saved locally. Ready to launch.";
  }

  openers.forEach((btn) => {
    btn.addEventListener("click", () => {
      openModal();
    });
  });

  if (backdrop) {
    backdrop.addEventListener("click", (event) => {
      if (event.target === backdrop) closeModal();
    });
  }

  if (closeBtn) {
    closeBtn.addEventListener("click", closeModal);
  }

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && backdrop?.classList.contains("is-visible")) {
      closeModal();
    }
  });

  accessForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const email = (emailInput?.value || "").trim();
    const isValid = /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email);

    if (!isValid) {
      if (messageEl) messageEl.textContent = "Enter a valid email to continue.";
      emailInput?.focus();
      return;
    }

    localStorage.setItem("teacherToolsEmail", email);
    savedEmail = email;
    if (messageEl) messageEl.textContent = "Saved locally. We will wire this to sign-in later.";
    closeModal();
    launchWorkspace();
  });

  function openModal() {
    if (!backdrop) return;
    backdrop.classList.add("is-visible");
    if (barFill) barFill.style.width = "0%";
    if (messageEl) messageEl.textContent = "";
    if (statusEl) {
      statusEl.textContent = savedEmail ? "Email saved locally. Ready to launch." : "Ready when you are.";
    }
    setTimeout(() => emailInput?.focus(), 120);
  }

  function closeModal() {
    backdrop?.classList.remove("is-visible");
  }

  function launchWorkspace() {
    if (isLaunching) return;
    isLaunching = true;
    homeShell.classList.add("launching");
    if (launchBtn) launchBtn.disabled = true;
    if (barFill) barFill.style.width = "100%";
    if (statusEl) statusEl.textContent = "Preparing your workspace...";

    setTimeout(() => {
      if (statusEl) statusEl.textContent = "Loading Teacher Tools...";
    }, 520);

    setTimeout(() => {
      homeShell.classList.add("fade-out");
      appRoot.classList.remove("is-hidden");
      window.scrollTo({ top: 0, behavior: "smooth" });
    }, 940);

    setTimeout(() => {
      homeShell.style.display = "none";
    }, 1400);
  }
})();
