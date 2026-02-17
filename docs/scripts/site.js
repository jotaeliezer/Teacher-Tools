(() => {
  const MAIL_TO = "support@teachertools.local";
  const FORM_ENDPOINT = "/api/contact"; // placeholder; replace with your backend endpoint

  const waitlistButtons = document.querySelectorAll("[data-mailto='waitlist']");
  waitlistButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      launchToast("Opening mail app for waitlist...", "info");
      openMail(
        MAIL_TO,
        "TeacherTools waitlist",
        "Hi TeacherTools team,%0D%0A%0D%0AI would like to join the waitlist for Pro/School.%0D%0A%0D%0AName:%0D%0ASchool:%0D%0ANeed:"
      );
    });
  });

  const contactForm = document.getElementById("contactForm");
  if (contactForm) {
    contactForm.addEventListener("submit", (event) => {
      event.preventDefault();
      const name = document.getElementById("contactName")?.value.trim() || "";
      const email = document.getElementById("contactEmail")?.value.trim() || "";
      const message = document.getElementById("contactMessage")?.value.trim() || "";
      if (!email) {
        launchToast("Enter your email so we can reply.", "warn");
        return;
      }

      // Attempt fetch first; fall back to mailto
      fetch(FORM_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, email, message }),
      })
        .then((res) => {
          if (!res.ok) throw new Error("Request failed");
          launchToast("Thanks! We received your message.", "success");
          contactForm.reset();
        })
        .catch(() => {
          const body = [
            "Hi TeacherTools team,",
            "",
            `Name: ${name || "(not provided)"}`,
            `Email: ${email || "(not provided)"}`,
            "",
            "Message:",
            message || "(empty)",
          ].join("\n");
          launchToast("Opening mail app as a fallback.", "info");
          openMail(MAIL_TO, "TeacherTools contact", encodeURIComponent(body));
        });
    });
  }

  const accordionItems = document.querySelectorAll(".accordion-item");
  accordionItems.forEach((item) => {
    const header = item.querySelector(".accordion-header");
    header?.addEventListener("click", () => {
      const isOpen = item.classList.contains("is-open");
      accordionItems.forEach((i) => i.classList.remove("is-open"));
      if (!isOpen) item.classList.add("is-open");
    });
  });

  const toolkitLink = document.querySelector('a[href="#homeFeatures"]');
  toolkitLink?.addEventListener("click", (event) => {
    event.preventDefault();
    const target = document.getElementById("homeFeatures");
    const nav = document.querySelector(".home-nav");
    if (!target) return;
    const navOffset = (nav?.offsetHeight || 0) + 24;
    const top = target.getBoundingClientRect().top + window.scrollY - navOffset;
    window.scrollTo({ top, behavior: "smooth" });
  });

  function openMail(to, subject, bodyEncoded) {
    const href = `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${bodyEncoded}`;
    window.location.href = href;
  }

  function launchToast(message, tone = "info") {
    let container = document.getElementById("toastContainer");
    if (!container) {
      container = document.createElement("div");
      container.id = "toastContainer";
      container.style.position = "fixed";
      container.style.bottom = "16px";
      container.style.right = "16px";
      container.style.display = "grid";
      container.style.gap = "8px";
      container.style.zIndex = "4000";
      document.body.appendChild(container);
    }
    const toast = document.createElement("div");
    toast.textContent = message;
    toast.style.padding = "10px 12px";
    toast.style.borderRadius = "10px";
    toast.style.boxShadow = "0 10px 24px rgba(0,0,0,.16)";
    toast.style.color = "#1b1b1b";
    toast.style.background = "#fff";
    toast.style.border = "1px solid rgba(27,27,27,.12)";
    if (tone === "success") toast.style.borderColor = "rgba(193,216,47,.8)";
    if (tone === "warn") toast.style.borderColor = "rgba(255,221,0,.8)";
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3200);
  }
})();
