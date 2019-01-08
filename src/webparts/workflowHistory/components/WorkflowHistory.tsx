import * as React from 'react';
import styles from './WorkflowHistory.module.scss';
import { IWorkflowHistoryProps } from './IWorkflowHistoryProps';
import { IWorkflowHistoryState } from './IWorkflowHistoryState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, IColumn, SelectionMode, Selection, } from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner } from 'office-ui-fabric-react/lib/components/Spinner';
import { CommandBar } from 'office-ui-fabric-react/lib/components/CommandBar';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/components/Dialog';
import { IContextualMenuItem, IContextualMenu } from "office-ui-fabric-react/lib/ContextualMenu";
import { TextField } from 'office-ui-fabric-react/lib/components/TextField';

import { PrimaryButton } from 'office-ui-fabric-react/lib/components/Button';
import { sp } from "@pnp/sp";
import wfItem from "./wfItem";
export default class WorkflowHistory extends React.Component<IWorkflowHistoryProps, IWorkflowHistoryState> {

  public constructor(props: IWorkflowHistoryProps) {
    super();
    debugger;
    this.state = { itemId: null, wfHistory: [] };
    console.log("in Construrctor");
    debugger;
  }
  public fetchItems(): void {
    sp.web.lists.getByTitle("Workflow History").items
      //.filter(`Item eq '${this.state.itemId}'`)
     // .top(200).get()
      .getAll()
      .then((items) => {
        let newItems: Array<wfItem> = [];
        debugger;
        for (let item of items) {
          if (item["Item"] === parseInt(this.state.itemId)) {
            let newItem: wfItem = {
              itemID: item["Item"],
              occurred: item["Occurred"],
              wfName: item["WorkflowAssociation"],
              message: item["Description"],
            };
            newItems.push(newItem)
          }
        }

        this.setState((current) => ({ ...current, wfHistory: newItems }));
      }).catch((e) => {
        console.log(e);
        debugger;
      });
  }

  public render(): React.ReactElement<IWorkflowHistoryProps> {
    return (
      <div className={styles.workflowHistory} >
        <TextField disabled={false} label="Item ID" value={this.state.itemId !== null ? this.state.itemId : ""}
          onChanged={(newValue) => {
            debugger;
            this.setState((current) => ({ ...current, itemId: newValue }));
          }
          }
        />
        <PrimaryButton label="Get Wofkflow History" onClick={this.fetchItems.bind(this)}>
        Get Wofkflow History
        </PrimaryButton>
        <DetailsList

          selectionMode={SelectionMode.none}
          items={this.state.wfHistory}
          columns={[
            { isResizable: true, minWidth: 40, maxWidth: 40, fieldName: "itemID", key: "itemID", name: "itemID" },

            { isResizable: true, minWidth: 110, maxWidth: 110, fieldName: "occurred", key: "occurred", name: "occurred" },
            { isResizable: true, minWidth: 175, maxWidth: 175, fieldName: "wfName", key: "wfName", name: "wfName" },
            {
              isResizable: true, minWidth: 250, fieldName: "message", key: "message", name: "message",
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return (
                  <TextField multiline={true} value={item.message} />
                );
              },
            }
          ]}
        />

      </div>
    );
  }
}
