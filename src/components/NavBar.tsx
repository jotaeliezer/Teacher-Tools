import { NavLink, Link } from "react-router-dom";

export default function NavBar() {
  return (
    <div className="home-nav">
      <div className="brand">
        <img
          src="/icon/teacher_tools_logo2.png"
          alt="TeacherTools logo"
          className="brand-logo"
        />
        <div className="brand-text">
          <strong>TeacherTools</strong>
          <span>Data-ready reporting</span>
        </div>
      </div>
      <div className="nav-links">
        <NavLink to="/">Home</NavLink>
        <NavLink to="/pricing">Pricing</NavLink>
        <NavLink to="/faq">FAQ</NavLink>
        <NavLink to="/contact">Contact</NavLink>
      </div>
      <span className="nav-beta">Local-only beta</span>
      <Link className="btn primary nav-launch" to="/teacher-tools">
        Launch TeacherTools
      </Link>
    </div>
  );
}
