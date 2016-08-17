import * as React from "react";

import {
  IExtSharingReportsProps,
  IGetExtContentFuncResponse,
  ControlMode,
  ISecurableObject
} from "./ExtSharingReportsInterfaces";

import {
  Logger
} from "./Logger";

import {
  FocusZone,
  FocusZoneDirection,
  KeyCodes,
  Spinner,
  SpinnerType,
  Label,
  css
} from "@ms/office-ui-fabric-react";

import styles from "./ExtSharingReports.module.scss";

export default class ExtContentTable extends React.Component<IExtSharingReportsProps, IGetExtContentFuncResponse> {
  private log: Logger;

  constructor() {
    super();
    this.log = new Logger("ExtContentTable");
  }

  public componentWillMount(): void {
    this.log.logInfo("componentWillMount");
    this._setStateToLoading();
  }

  public componentDidMount(): void {
    this.log.logInfo("componentDidMount");
    this._updateState();
  }

  private updatedOnce: boolean = false;

  public componentDidUpdate(): void {
    this.log.logInfo("componentDidUpdate");
    if (this.state.timeStamp === this.props.store.timeStamp) {
      // Do nothing, as the data will be the same
    } else {
      if (!this.updatedOnce) {
        this.updatedOnce = true;
        this._updateState();
      }
    }
  }

  private _updateState(): void {
    this.log.logInfo("_updateState");
    this._setStateToLoading();
    this.props.store.getAllExtDocuments()
    .then((r) => {
      this.setState(r);
      this.log.logInfo("_setStateToContent");
    });
  }

  private _setStateToLoading(): void {
    this.log.logInfo("_setStateToLoading");
    this.setState({
      extContent: [],
      controlMode: ControlMode.Loading,
      message: "Working on it...",
      timeStamp: (new Date()).getTime()
    });
  }

  public render(): JSX.Element {
    this.log.logInfo("render");
    if (this.state && this.state.controlMode === ControlMode.Loading) {
      return (
        <Spinner type={ SpinnerType.large } label={this.state.message} />
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Message) {
      return (
        <Label>{this.state.message}</Label>
      );
    }
    else if (this.state && this.state.controlMode === ControlMode.Content) {
      // TODO: Extract Table and TableRow into their own classes
      const tableCellClasses: string = css(styles.msTableCellNoWrap, "ms-Table-cell");
      const TableRow = (row: ISecurableObject) => (
        <tr className="ms-Table-row">
          <td className={tableCellClasses}>{row.Type}</td>
          <td className={tableCellClasses}>{row.Title}</td>
          <td className={tableCellClasses}>{row.LastModifiedTime}</td>
          <td className={tableCellClasses}>{row.SharedWith}</td>
          <td className={tableCellClasses}>{row.SharedBy}</td>
        </tr>
      );
      const tableRows: JSX.Element[] = this.state.extContent.map(c => {
        return (
          <TableRow
            key={c.URL}
            Type={c.Type}
            Title={c.Title}
            URL={c.URL}
            FileExtension={c.FileExtension}
            SharedBy={c.SharedBy}
            SharedWith={c.SharedWith}
            LastModifiedTime={c.LastModifiedTime}
            />
        );
      });
      const Table = () => (
        <div className={styles.msTableOverflow}>
          <table className="ms-Table">
              <tr className="ms-Table-row">
                <td className={tableCellClasses}>Type</td>
                <td className={tableCellClasses}>Title</td>
                <td className={tableCellClasses}>Modified</td>
                <td className={tableCellClasses}>Shared With</td>
                <td className={tableCellClasses}>Shared By</td>
              </tr>
            <tbody>
              {tableRows}
            </tbody>
          </table>
        </div>
      );

      return (
        <FocusZone
          direction={ FocusZoneDirection.vertical }
          isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }>
            <Table>
            </Table>
        </FocusZone>
      );
    }
    else {
      this.log.logError(`ControlMode is not supported ${this.state.controlMode}`);
    }
  }
}
