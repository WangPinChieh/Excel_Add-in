import * as React from "react";
import Header from "./Header";
import Progress from "./Progress";
import Home from "./Home";

/* global require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App = (props: AppProps) => {
  const { title, isOfficeInitialized } = props;

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
      <Home />
    </div>
  );
};

export default App;
