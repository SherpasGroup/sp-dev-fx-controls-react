import * as React from 'react';
import { ITimeComponentProps } from './ITimeComponentProps';
import {
  Dropdown,
  IDropdownOption
} from 'office-ui-fabric-react/lib/components/Dropdown';

/**
 * Minutes component, renders the minutes dropdown
 */
export default class MinutesDropdownComponent extends React.Component<ITimeComponentProps, {}> {
  private _minutes: IDropdownOption[];

  constructor(props: ITimeComponentProps) {
    super(props);
    this._initMinutesOptions();
  }

  public render(): JSX.Element {
    return (
      <Dropdown
        disabled={this.props.disabled}
        label=""
        options={this._minutes}
        selectedKey={this.props.value}
        onChanged={this.handleChange}
      />
    );
  }

  private handleChange = (option: IDropdownOption): void => {
    this.props.onChange(option.key as number);
  }

  private _initMinutesOptions() {
    const minutes: IDropdownOption[] = [];
    for (let j = 0; j < 60; j++) {
      const digitMin: string = j < 10 ? `0${j}` : j.toString();
      minutes.push({ key: j, text: digitMin });
    }
    this._minutes = minutes;
  }
}
