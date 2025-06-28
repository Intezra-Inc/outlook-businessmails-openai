import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Office, module, require */

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <FluentProvider theme={webLightTheme}>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </FluentProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
