import * as React from "react";

import styles from "./ExtSharingReports.module.scss";

import {
  IExtSharingReportsProps
} from "./ExtSharingReportsInterfaces";

export default class ExtSharingReports extends React.Component<IExtSharingReportsProps, {}> {
  public render(): JSX.Element {
    return (
      <div className={styles.extSharingReports}>
        <div>This is the <b>I'll get some content later...</b> webpart.</div>
      </div>
    );
  }
}
