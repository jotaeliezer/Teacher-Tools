import { useEffect, useRef, useState } from "react";
import { Link } from "react-router-dom";

type TeacherToolsProps = {
  isActive: boolean;
};

const SCRIPT_QUEUE = [
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
  "https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js",
  "/scripts/state.js",
  "/scripts/main.js",
  "/scripts/comments.js",
  "/scripts/animations.js",
];

function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-tt-src="${src}"]`);
    if (existing) {
      resolve();
      return;
    }
    const script = document.createElement("script");
    script.src = src;
    script.async = false;
    script.dataset.ttSrc = src;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.body.appendChild(script);
  });
}

function rewriteLegacyLinks(container: HTMLElement) {
  const map: Record<string, string> = {
    "home.html": "/",
    "pricing.html": "/pricing",
    "faq.html": "/faq",
    "contact.html": "/contact",
  };
  container.querySelectorAll<HTMLAnchorElement>("a[href]").forEach((anchor) => {
    const href = anchor.getAttribute("href") || "";
    const next = map[href];
    if (next) anchor.setAttribute("href", next);
  });

  container.querySelectorAll<HTMLImageElement>("img[src]").forEach((img) => {
    const src = img.getAttribute("src") || "";
    if (src.startsWith("icon/")) {
      img.setAttribute("src", `/${src}`);
    }
  });
}

export default function TeacherTools({ isActive }: TeacherToolsProps) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [ready, setReady] = useState(false);
  const loadStarted = useRef(false);

  useEffect(() => {
    if (!isActive || loadStarted.current) return;
    loadStarted.current = true;

    const run = async () => {
      const response = await fetch("/teacher_tools.html", { cache: "no-cache" });
      const html = await response.text();
      const doc = new DOMParser().parseFromString(html, "text/html");
      doc.querySelectorAll("script").forEach((script) => script.remove());
      if (containerRef.current) {
        containerRef.current.innerHTML = doc.body.innerHTML;
        rewriteLegacyLinks(containerRef.current);
      }

      for (const src of SCRIPT_QUEUE) {
        await loadScript(src);
      }
      setReady(true);
    };

    run().catch((error) => {
      console.error("TeacherTools init failed", error);
    });
  }, [isActive]);

  return (
    <div className={`teacher-tools-shell ${isActive ? "" : "is-hidden"}`}>
      <div className="app-embed-bar">
        <Link className="btn btn-ghost" to="/">
          Back to Home
        </Link>
        <span className="muted">
          {ready ? "TeacherTools ready" : "Preparing TeacherTools..."}
        </span>
      </div>
      <div ref={containerRef} />
    </div>
  );
}
