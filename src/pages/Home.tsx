import { Link } from "react-router-dom";
import PhaserHero from "../components/PhaserHero.tsx";

export default function Home() {
  const handleScroll = (event: React.MouseEvent<HTMLAnchorElement>) => {
    event.preventDefault();
    const target = document.getElementById("homeFeatures");
    const nav = document.querySelector<HTMLElement>(".home-nav");
    if (!target) return;
    const navOffset = (nav?.offsetHeight || 0) + 24;
    const top = target.getBoundingClientRect().top + window.scrollY - navOffset;
    window.scrollTo({ top, behavior: "smooth" });
  };

  return (
    <>
      <section className="full-section hero-section band-light">
        <div className="hero-ambient">
          <PhaserHero />
        </div>
        <div className="home-layout">
          <div className="home-hero">
            <div className="home-meta">
              <span className="pill">TeacherTools</span>
              <span className="pill pill-ghost">Spreadsheet-ready</span>
              <span className="pill pill-ghost">Report-friendly</span>
            </div>
            <h1>All-in-one tools for fast, consistent classroom reporting.</h1>
            <p className="lede">
              Upload rosters and marks, tidy headers, filter, and export print-ready sheets without juggling tabs.
              Jump into TeacherTools and keep columns, comments, and classes aligned.
            </p>
            <div className="whats-new">
              <div className="pill pill-ghost">What's new</div>
              <div className="whats-new-items">
                <div className="whats-new-item">
                  <span className="stat-label">Print view</span>
                  <span className="stat-value">Refreshed headers</span>
                </div>
                <div className="whats-new-item">
                  <span className="stat-label">Comments</span>
                  <span className="stat-value">Token helpers</span>
                </div>
                <div className="whats-new-item">
                  <span className="stat-label">Data</span>
                  <span className="stat-value">Folder upload</span>
                </div>
              </div>
            </div>
            <div className="hero-actions">
              <Link className="btn primary launch-btn" to="/teacher-tools">
                Launch TeacherTools
              </Link>
              <a className="btn btn-ghost" href="#homeFeatures" onClick={handleScroll}>
                See the toolkit
              </a>
            </div>
          </div>
        </div>
      </section>

      <section className="full-section features-section band-muted" id="homeFeatures">
        <div className="section-heading">
          <h2>Built for fast, consistent reporting</h2>
          <p className="section-lede">
            Drop in CSV/XLSX files, tidy headers, and export without jumping between tabs.
          </p>
        </div>
        <div className="home-grid wide">
          <div className="home-card fade-in delay-1">
            <h3>Built for teachers</h3>
            <p>Drag in CSV/XLSX files, keep columns aligned, and jump between classes without losing context.</p>
            <ul className="feature-list">
              <li>Folder upload with quick file switching</li>
              <li>Persisted header renames and sorting</li>
              <li>Undo/redo for safer edits</li>
            </ul>
          </div>
          <div className="home-card fade-in delay-2">
            <h3>Instant exports</h3>
            <p>Send clean CSV/XLSX to colleagues or print formatted reports tailored to each class.</p>
            <ul className="feature-list">
              <li>Print view with templates</li>
              <li>Mark highlighting and filters</li>
              <li>CSV/XLSX export of visible columns</li>
            </ul>
          </div>
          <div className="home-card fade-in delay-3">
            <h3>Comments that scale</h3>
            <p>Use the report comment builder to combine marks, pronouns, and reusable statements in seconds.</p>
            <ul className="feature-list">
              <li>Student-by-student previews</li>
              <li>Assignment token support</li>
              <li>Copy-ready report text</li>
            </ul>
          </div>
        </div>
      </section>

      <section className="full-section workflow-section band-light" id="homeWorkflow">
        <div className="section-heading">
          <h2>Upload, align, and share</h2>
          <p className="section-lede">
            Keep filters, hidden columns, comments, and exports in sync across your classes.
          </p>
        </div>
        <div className="home-panels">
          <div className="home-panel">
            <div className="panel-eyebrow">Workflow</div>
            <h3>Upload, align, export</h3>
            <p>Load multiple CSV/XLSX files, align headers once, and keep filters, hidden columns, and comments in sync.</p>
          </div>
          <div className="home-panel">
            <div className="panel-eyebrow">Reporting</div>
            <h3>Comment builder that respects your data</h3>
            <p>Merge pronouns, marks, and templates into clean report text. Copy per-student or everything at once.</p>
            <div className="panel-stats">
              <div className="stat">
                <span className="stat-label">Classes</span>
                <span className="stat-value">Multi-class ready</span>
              </div>
              <div className="stat">
                <span className="stat-label">Exports</span>
                <span className="stat-value">CSV &amp; XLSX</span>
              </div>
              <div className="stat">
                <span className="stat-label">Modes</span>
                <span className="stat-value">Light &amp; Dark</span>
              </div>
            </div>
          </div>
        </div>
      </section>

      <section className="full-section testimonials-section band-muted">
        <div className="section-heading">
          <h2>What's working in real classrooms</h2>
          <p className="section-lede">
            Feedback from teams using TeacherTools for rosters, marks, and report comments.
          </p>
        </div>
        <div className="section-alt">
          <div className="testimonials">
            <div className="testimonial-card">
              <div className="testimonial-avatar">JR</div>
              <div className="testimonial-body">
                <p className="testimonial-quote">"I stopped juggling five tabs to print reports."</p>
                <div className="testimonial-meta">Grade 7 math</div>
              </div>
            </div>
            <div className="testimonial-card">
              <div className="testimonial-avatar">SL</div>
              <div className="testimonial-body">
                <p className="testimonial-quote">"Folder upload and header rename saved me hours."</p>
                <div className="testimonial-meta">Science lead</div>
              </div>
            </div>
            <div className="testimonial-card">
              <div className="testimonial-avatar">PD</div>
              <div className="testimonial-body">
                <p className="testimonial-quote">"Comments stay consistent between terms."</p>
                <div className="testimonial-meta">Program director</div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <section className="full-section privacy-section band-light" id="homePrivacy">
        <div className="home-layout">
          <div className="home-aside" style={{ width: "100%" }}>
            <div className="home-mini-card">
              <h4>What happens to my data?</h4>
              <p>
                Files are processed locally in your browser. When we hook up email-based access, this form will
                route to a storage service you approve.
              </p>
            </div>
            <div className="home-mini-card">
              <h4>Need to dive straight in?</h4>
              <p>Hit "Launch TeacherTools" and you will jump into the app after a short launch animation.</p>
            </div>
          </div>
        </div>
      </section>
    </>
  );
}
