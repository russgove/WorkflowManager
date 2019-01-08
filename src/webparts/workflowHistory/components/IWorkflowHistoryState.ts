import wfItem from "./wfItem";
export interface IWorkflowHistoryState {
  itemId: string;
  wfHistory: Array<wfItem>;
}
