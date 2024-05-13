import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { PrimaryButton, DefaultButton, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { DialogFooter, DialogContent } from 'office-ui-fabric-react';

export interface IHtmlDialogContentProps {
  title?: string;
  message?: string;
  closeText?: string;
  sendText?: string;
  send: () => Promise<void>;
  close: () => void;
  linkUrl?: string;
  linkTitle?: string;
  disable?: boolean;
  buttonStyle?: {};
  textAreaStyle?: {};
  textAreaValue?: string;
  invalidMaterial?: string;
  invalidTextAreaStyle?: {};
  readTextArea?: boolean;
  readInvalidTextArea?: boolean;
  isLoading?: boolean;
  updateText?: string;
}

class HtmlDialogContent extends React.Component<IHtmlDialogContentProps, {}> {
  constructor(props: any) {
    super(props);
  }

  public render(): JSX.Element {
    return <DialogContent
      title={this.props.title}
      subText={this.props.message}
      onDismiss={this.props.close}
      showCloseButton={true}
    >
      <DialogFooter>
        <DefaultButton text={this.props.closeText} title={this.props.closeText} onClick={() => { this.props.close(); }} />
        <PrimaryButton text={this.props.sendText} title={this.props.sendText} style={this.props.buttonStyle} onClick={() => { this.props.send(); }} />
      </DialogFooter>
    </DialogContent>;
  }
}

export default class HtmlPickerDialog extends BaseDialog {
  constructor(private send: () => Promise<void>, private dialogTitle?: string, private dialogMessage?: string, private closeButtonText?: string, private sendButtonText?: string, private linkUrl?: string, private linkTitle?: string, private disable?: boolean, private buttonStyle?: {}, private textAreaStyle?: {}, private textAreaValue?: string, private invalidMaterial?: string, private invalidTextAreaStyle?: {}, private readTextArea?: boolean, private readInvalidTextArea?: boolean, private isLoading?: boolean) {
    super();
  }

  public render(): void {
    ReactDOM.render(<HtmlDialogContent
      title={this.dialogTitle}
      message={this.dialogMessage}
      linkUrl={this.linkUrl}
      linkTitle={this.linkTitle}
      closeText={this.closeButtonText}
      sendText={this.sendButtonText}
      send={this.send}
      close={this.close}
      disable={this.disable}
      buttonStyle={this.buttonStyle}
      textAreaStyle={this.textAreaStyle}
      textAreaValue={this.textAreaValue}
      invalidMaterial={this.invalidMaterial}
      invalidTextAreaStyle={this.invalidTextAreaStyle}
      readTextArea={this.readTextArea}
      readInvalidTextArea={this.readInvalidTextArea}
      isLoading={this.isLoading}
    />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: true
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

}
