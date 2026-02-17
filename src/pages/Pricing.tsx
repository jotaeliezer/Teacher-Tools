import { Link } from "react-router-dom";
import { launchToast, openMail } from "../utils/site.ts";

const MAIL_TO = "support@teachertools.local";

export default function Pricing() {
  const handleWaitlist = () => {
    launchToast("Opening mail app for waitlist...", "info");
    openMail(
      MAIL_TO,
      "TeacherTools waitlist",
      "Hi TeacherTools team,%0D%0A%0D%0AI would like to join the waitlist for Pro/School.%0D%0A%0D%0AName:%0D%0ASchool:%0D%0ANeed:"
    );
  };

  return (
    <section className="full-section band-light">
      <div className="section-heading page-hero">
        <h1>Simple pricing for classrooms and teams.</h1>
        <p className="section-lede">
          Start free while everything runs locally. Move to Pro or School when sync and sign-in arrive.
        </p>
      </div>
      <div className="home-grid">
        <div className="home-card fade-in delay-1">
          <h3>Free</h3>
          <p>Everything runs in your browser. Upload CSV/XLSX, clean columns, and export without limits.</p>
          <ul className="feature-list">
            <li>Unlimited file uploads</li>
            <li>Print view and exports</li>
            <li>Comment builder</li>
          </ul>
          <div className="hero-actions">
            <Link className="btn primary" to="/teacher-tools">
              Use Free
            </Link>
          </div>
        </div>
        <div className="home-card fade-in delay-2">
          <h3>Pro (coming soon)</h3>
          <p>Account-based access with cloud save, shared templates, and priority support.</p>
          <ul className="feature-list">
            <li>Secure sign-in</li>
            <li>Team template sharing</li>
            <li>Support and onboarding</li>
          </ul>
          <div className="hero-actions">
            <button className="btn primary" type="button" onClick={handleWaitlist}>
              Join waitlist
            </button>
          </div>
        </div>
        <div className="home-card fade-in delay-3">
          <h3>School</h3>
          <p>Site licenses for multiple teachers with centralized settings and deployment help.</p>
          <ul className="feature-list">
            <li>Centralized admin</li>
            <li>Custom domains</li>
            <li>Training sessions</li>
          </ul>
          <div className="hero-actions">
            <Link className="btn primary" to="/contact">
              Talk to us
            </Link>
          </div>
        </div>
      </div>
      <div className="home-grid" style={{ marginTop: "18px" }}>
        <div className="home-card">
          <h4>Billing transparency</h4>
          <p>TeacherTools is free while it is local-only. Paid tiers appear once sync and sign-in ship.</p>
        </div>
        <div className="home-card">
          <h4>Need a purchase order?</h4>
          <p>Email us for invoicing, purchase orders, or multi-seat quotes.</p>
        </div>
      </div>
    </section>
  );
}
