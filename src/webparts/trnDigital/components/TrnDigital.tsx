import * as React from "react";
// import styles from "./TrnDigital.module.scss";
import { ITrnDigitalProps } from "./ITrnDigitalProps";
import { sp } from "@pnp/sp/presets/all";
import { escape } from "@microsoft/sp-lodash-subset";
import MainComponent from "./MainComponent";
import "./style.css";
export default class TrnDigital extends React.Component<ITrnDigitalProps, {}> {
  constructor(prop: ITrnDigitalProps, state: {}) {
    super(prop);
    sp.setup({ spfxContext: this.props.context });
  }
  public render(): React.ReactElement<ITrnDigitalProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <MainComponent
        spcontext={this.props.context}
        libraryName={this.props.libraryName}
      />
    );
  }
}
