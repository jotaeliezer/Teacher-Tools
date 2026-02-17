import { useState } from "react";
import { Link } from "react-router-dom";
import { launchToast, openMail } from "../utils/site.ts";

const MAIL_TO = "support@teachertools.local";
const FORM_ENDPOINT = "/api/contact";

export default function Contact() {
  const [form, setForm] = useState({ name: "", email: "", message: "" });

  const update = (field: "name" | "email" | "message") => (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    setForm((prev) => ({ ...prev, [field]: event.target.value }));
  };

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!form.email.trim()) {
      launchToast("Enter your email so we can reply.", "warn");
      return;
    }

    try {
      const response = await fetch(FORM_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      if (!response.ok) throw new Error("Request failed");
      launchToast("Thanks! We received your message.", "success");
      setForm({ name: "", email: "", message: "" });
    } catch (error) {
      const body = [
        "Hi TeacherTools team,",
        "",
        `Name: ${form.name || "(not provided)"}`,
        `Email: ${form.email || "(not provided)"}`,
        "",
        "Message:",
        form.message || "(empty)",
      ].join("\n");
      launchToast("Opening mail app as a fallback.", "info");
      openMail(MAIL_TO, "TeacherTools contact", encodeURIComponent(body));
    }
  };

  return (
    <section className="full-section band-light">
      <div className="section-heading page-hero">
        <h1>Tell us what you need.</h1>
        <p className="section-lede">
          Questions, feature ideas, or school rollouts - send a note and we will reply within a school day.
        </p>
      </div>
      <div className="home-grid">
        <div className="home-card">
          <h3>Support</h3>
          <p>Need help with uploads, printing, or comments?</p>
          <ul className="feature-list">
            <li>Email: support@teachertools.local</li>
            <li>Response target: 1 business day</li>
            <li>Timezone: EST</li>
          </ul>
        </div>
        <div className="home-card">
          <h3>Schools</h3>
          <p>Rolling out to multiple teachers? We can help with onboarding and training.</p>
          <ul className="feature-list">
            <li>Multi-seat planning</li>
            <li>Training sessions</li>
            <li>Purchase orders</li>
          </ul>
        </div>
        <div className="home-card">
          <h3>Features</h3>
          <p>Want a specific workflow or integration? Send requests our way.</p>
          <ul className="feature-list">
            <li>Template requests</li>
            <li>Export formats</li>
            <li>Integrations</li>
          </ul>
        </div>
      </div>
      <div className="home-panel contact-panel" style={{ marginTop: "18px" }}>
        <h3>Drop a line</h3>
        <p>Send your name, school, and what you need. We will follow up with next steps.</p>
        <form id="contactForm" onSubmit={handleSubmit}>
          <div className="access-row" style={{ marginTop: "10px" }}>
            <input
              id="contactName"
              type="text"
              placeholder="Your name"
              value={form.name}
              onChange={update("name")}
            />
            <input
              id="contactEmail"
              type="email"
              placeholder="you@school.edu"
              autoComplete="email"
              value={form.email}
              onChange={update("email")}
            />
          </div>
          <div className="access-row" style={{ marginTop: "8px" }}>
            <textarea
              id="contactMessage"
              style={{
                width: "100%",
                minHeight: "120px",
                padding: "12px",
                borderRadius: "12px",
                border: "1px solid #d5d5d5",
              }}
              placeholder="How can we help?"
              value={form.message}
              onChange={update("message")}
            />
          </div>
          <div className="hero-actions" style={{ marginTop: "10px" }}>
            <button className="btn primary" id="contactSubmit" type="submit">
              Send message
            </button>
            <Link className="btn btn-ghost" to="/teacher-tools">
              Open TeacherTools
            </Link>
          </div>
        </form>
      </div>
      <div className="home-grid" style={{ marginTop: "18px" }}>
        <div className="home-card">
          <h4>Office hours</h4>
          <p>Weekdays 9a-5p EST. We often reply sooner for urgent issues.</p>
        </div>
        <div className="home-card">
          <h4>Best way to reach us</h4>
          <p>Email first; if you prefer a call, mention times that work and we will schedule.</p>
        </div>
      </div>
    </section>
  );
}
