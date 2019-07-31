import * as React from 'react';
import { ITimeComponentProps } from './ITimeComponentProps';
import { MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { TimeHelper } from './TimeHelper';

/**
 * Seconds component, renders the seconds dropdown
 */
export default class SecondsComponent extends React.Component<ITimeComponentProps, {}> {

  public render(): JSX.Element {
    return (
      <MaskedTextField disabled={this.props.disabled}
        label=""
        value={this.props.value ? TimeHelper.prefixZero(this.props.value.toString()) : "00"}
        onGetErrorMessage={this.handleGetErrorMessage}
        mask="59"
        maskFormat={{
          '5': /[0-5]/,
          '9': /[0-9]/
        }}
      />
    );
  }

  private handleGetErrorMessage = (value: string): string => {
    const seconds: number = parseInt(value.length > 2 ? value.substring(0, 2) : value);
    this.props.onChange(seconds);
    return '';
  }
}
