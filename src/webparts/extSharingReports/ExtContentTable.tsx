import * as React from "react";

import styles from "./ExtSharingReports.module.scss";

import {
  IExtSharingReportsProps
} from "./ExtSharingReportsInterfaces";

export default class ExtContentTable extends React.Component<IExtSharingReportsProps, {}> {
  public render(): JSX.Element {
    return (
      <div className={styles.extSharingReports}>
        <div>Hello world</div>
      </div>
    );
  }
}
