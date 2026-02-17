export function openMail(to, subject, bodyEncoded) {
  const href = `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${bodyEncoded}`;
  window.location.href = href;
}

export function launchToast(message, tone = "info") {
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
