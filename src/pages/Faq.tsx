import { useState } from "react";
import { Link } from "react-router-dom";

const ITEMS = [
  {
    title: "Where does my data live?",
    body: "It stays in your browser. Nothing leaves your machine unless you export a file.",
  },
  {
    title: "What files can I upload?",
    body: "CSV, XLSX, and XLS. You can drag multiple files or load a folder.",
  },
  {
    title: "Do I need an account?",
    body: "Not yet. We ask for email now to prep for future sign-in; it is only stored locally.",
  },
  {
    title: "Can I print formatted reports?",
    body: "Yes. Use Print View to format attendance, marking, and drill sheets before printing or exporting.",
  },
  {
    title: "How do comments work?",
    body: "Use the builder to merge pronouns, marks, and templates. Copy per student or all at once.",
  },
  {
    title: "How do I get help?",
    body: "Visit the contact page for questions, onboarding, or feature requests.",
  },
];

export default function Faq() {
  const [openIndex, setOpenIndex] = useState(0);

  return (
    <section className="full-section band-light">
      <div className="section-heading page-hero">
        <h1>Quick answers about TeacherTools.</h1>
        <p className="section-lede">Everything runs locally today. Email-based sign-in and sync will arrive soon.</p>
      </div>
      <div className="accordion" id="faqAccordion">
        {ITEMS.map((item, idx) => {
          const isOpen = openIndex === idx;
          return (
            <div key={item.title} className={`accordion-item ${isOpen ? "is-open" : ""}`}>
              <div
                className="accordion-header"
                role="button"
                tabIndex={0}
                onClick={() => setOpenIndex(isOpen ? -1 : idx)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    setOpenIndex(isOpen ? -1 : idx);
                  }
                }}
              >
                <p className="accordion-title">{item.title}</p>
                <div className="accordion-icon">+</div>
              </div>
              <div className="accordion-body">{item.body}</div>
            </div>
          );
        })}
      </div>
      <div className="home-grid" style={{ marginTop: "16px" }}>
        <div className="home-card">
          <h4>Still stuck?</h4>
          <p>Visit the contact page and tell us what you are trying to do. We usually respond within a school day.</p>
          <Link className="btn primary" to="/contact">
            Get help
          </Link>
        </div>
      </div>
    </section>
  );
}
