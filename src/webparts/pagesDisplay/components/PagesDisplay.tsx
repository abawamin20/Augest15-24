import * as React from "react";
import type { IPagesDisplayProps } from "./IPagesDisplayProps";
import PagesList from "./PagesList/PagesList";
import { SPComponentLoader } from "@microsoft/sp-loader";

export default class PagesDisplay extends React.Component<
  IPagesDisplayProps,
  {}
> {
  constructor(components: IPagesDisplayProps) {
    super(components);
    const cssURL =
      "https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);
    SPComponentLoader.loadCss(
      "https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css"
    );
  }
  public render(): React.ReactElement<IPagesDisplayProps> {
    return (
      <PagesList
        context={this.props.context}
        selectedViewId={this.props.selectedViewId}
        feedbackPageUrl={this.props.feedbackPageUrl}
      />
    );
  }
}
