import { Routes, Route, useLocation } from "react-router-dom";
import { useEffect, useState } from "react";
import MarketingLayout from "./components/MarketingLayout.tsx";
import Home from "./pages/Home.tsx";
import Pricing from "./pages/Pricing.tsx";
import Faq from "./pages/Faq.tsx";
import Contact from "./pages/Contact.tsx";
import TeacherTools from "./pages/TeacherTools.tsx";

export default function App() {
  const location = useLocation();
  const [teacherToolsReady, setTeacherToolsReady] = useState(false);
  const isTeacherToolsRoute = location.pathname === "/teacher-tools";

  useEffect(() => {
    if (isTeacherToolsRoute) setTeacherToolsReady(true);
  }, [isTeacherToolsRoute]);

  return (
    <>
      <Routes>
        <Route
          path="/"
          element={
            <MarketingLayout>
              <Home />
            </MarketingLayout>
          }
        />
        <Route
          path="/pricing"
          element={
            <MarketingLayout>
              <Pricing />
            </MarketingLayout>
          }
        />
        <Route
          path="/faq"
          element={
            <MarketingLayout>
              <Faq />
            </MarketingLayout>
          }
        />
        <Route
          path="/contact"
          element={
            <MarketingLayout>
              <Contact />
            </MarketingLayout>
          }
        />
      </Routes>
      {teacherToolsReady ? <TeacherTools isActive={isTeacherToolsRoute} /> : null}
    </>
  );
}
