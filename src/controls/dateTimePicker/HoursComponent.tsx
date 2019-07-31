import * as strings from 'ControlStrings';
import * as React from 'react';
import { IHoursComponentProps } from './ITimeComponentProps';
import { TimeConvention } from './DateTimeConventions';
import { MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { TimeHelper } from './TimeHelper';

/**
 * Hours component, this renders the hours dropdown
 */
export default class HoursComponent extends React.Component<IHoursComponentProps, {}> {

  public render(): JSX.Element {
    return (
      <MaskedTextField disabled={this.props.disabled}
        label=""
        value={this.props.value ? TimeHelper.hoursValue(this.props.value, this.props.timeConvention) : `${this.props.timeConvention === TimeConvention.Hours24 ? "00" : "12 AM"}`}
        mask={this.props.timeConvention === TimeConvention.Hours24 ? "29" : "19 AM"}
        maskFormat={{
          '1': /[0-1]/,
          '2': /[0-2]/,
          '9': /[0-9]/,
          'A': /[AaPp]/,
          'M': /[Mm]/
        }}
        onGetErrorMessage={this.handleGetErrorMessage}
      />
    );
  }

  private handleGetErrorMessage = (value: string): string => {
    let message = "";
    const hoursSplit = value.split(" ");
    const hoursValue = hoursSplit[0].length > 2 ? hoursSplit[0].substring(0, 2) : hoursSplit[0];
    let hours: number = parseInt(hoursValue);
    if (isNaN(hours)) {
      return strings.DateTimePickerHourValueInvalid;
    }

    if (this.props.timeConvention !== TimeConvention.Hours24) {
      if (!hoursSplit[1]) {
        return strings.DateTimePickerHourValueInvalid;
      }
      if (hoursSplit[1].toLowerCase().indexOf("pm") !== -1) {
        hours += 12;
        if (hours === 24) {
          hours = 12;
        }
      } else {
        if (hours === 12) {
          hours = 0;
        }
      }
    }

    if (hours < 0 || hours > 23) {
      return strings.DateTimePickerHourValueInvalid;
    }

    this.props.onChange(hours);
    return '';
  }
}
