import * as strings from 'ControlStrings';
import * as React from 'react';
import { IHoursComponentProps } from './ITimeComponentProps';
import { TimeConvention } from './DateTimeConventions';
import {
  Dropdown,
  IDropdownOption
} from 'office-ui-fabric-react/lib/components/Dropdown';
// import { TimeHelper } from './TimeHelper';

/**
 * Hours dropdown component, this renders the hours dropdown
 */
export default class HoursDropdownComponent extends React.Component<IHoursComponentProps, {}> {
  private _hours: IDropdownOption[];

  constructor(props: IHoursComponentProps) {
    super(props);
    this._initHoursOptions('AM', 'PM'); //props.amDesignator, props.pmDesignator);
  }

  public render(): JSX.Element {
    return (
      <Dropdown
        disabled={this.props.disabled}
        label=""
        options={this._hours}
        selectedKey={this.props.value}
        onChanged={this.handleChange}
        dropdownWidth={110}
      />
    );
  }

  private handleChange = (option: IDropdownOption): void => {
    this.props.onChange(option.key as number);
  }

  private _initHoursOptions(amDesignator: string, pmDesignator: string) {
    const hours: IDropdownOption[] = [];
    for (let i = 0; i < 24; i++) {
      let digit: string;
      if (this.props.timeConvention === TimeConvention.Hours24) {
        // 24 hours time convention
        if (i < 10) {
          digit = `0${i}`;
        } else {
          digit = i.toString();
        }
      } else {
        // 12 hours time convention
        if (i === 0) {
          digit = `12 ${amDesignator}`;
        } else if (i < 12) {
          digit = `${i} ${amDesignator}`;
        } else {
          if (i === 12) {
            digit = `12 ${pmDesignator}`;
          } else {
            digit = `${(i % 12)} ${pmDesignator}`;
          }
        }
      }
      hours.push({ key: i, text: digit });
    }
    this._hours = hours;
  }
}
