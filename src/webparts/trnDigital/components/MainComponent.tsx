import * as React from "react";
import Dashboard from "./Dashboard/Dashboard";

export default function MainComponent(props: any): JSX.Element {
  return (
    <div>
      <Dashboard spcontext={props.spcontext} libraryName={props.libraryName} />
    </div>
  );
}
