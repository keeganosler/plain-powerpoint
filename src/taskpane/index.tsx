import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { BrandVariants, createLightTheme, FluentProvider, Theme, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

const myNewTheme: BrandVariants = { 
  10: "#060201",
  20: "#231309",
  30: "#3C1C10",
  40: "#502414",
  50: "#652B16",
  60: "#7B3319",
  70: "#913A1C",
  80: "#A8421E",
  90: "#BB4E27",
  100: "#C5613C",
  110: "#CF7451",
  120: "#D78667",
  130: "#E0987D",
  140: "#E7AA93",
  150: "#EEBDAA",
  160: "#F4CFC1"
};

 const pptTheme: Theme = {
   ...createLightTheme(myNewTheme), 
};

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={pptTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
