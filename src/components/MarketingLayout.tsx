import NavBar from "./NavBar.tsx";
import Footer from "./Footer.tsx";
import type { PropsWithChildren } from "react";

export default function MarketingLayout({ children }: PropsWithChildren) {
  return (
    <>
      <div className="home-shell">
        <NavBar />
        {children}
      </div>
      <Footer />
    </>
  );
}
