import * as React from 'react';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { cloneDeep } from '@microsoft/sp-lodash-subset';



import { PrimaryButton, DefaultButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Autofill } from 'office-ui-fabric-react/lib/components/Autofill/Autofill';
import { Label } from 'office-ui-fabric-react/lib/Label';

import FieldErrorMessage from '../errorMessage/ErrorMessage';


import { sp } from '@pnp/sp';
import { ITermInfo, ITermSetInfo, ITermSet, ITerm } from '@pnp/sp/taxonomy';
import { SPTaxonomyService } from '../../services/SPTaxonomyService';
import styles from './ModernTaxonomyPicker.module.scss';
import * as strings from 'ControlStrings';

// TODO: this import should be replaced with our own tag-picker
import TermPicker from '../taxonomyPicker/TermPicker';
import { TaxonomyForm } from './taxonomyForm';
import { Guid } from '@microsoft/sp-core-library';
import { IIconProps, IStackTokens, ITag, Stack, TagPicker } from 'office-ui-fabric-react';

// TODO: remove/replace interface IPickerTerm
export interface IPickerTerm {
  name: string;
  key: string;
  path: string;
  termSet: string;
  termSetName?: string;
}

// TODO: remove/replace interface IPickerTerms
export interface IPickerTerms extends Array<IPickerTerm> { }

export interface IModernTaxonomyPickerProps {
  allowMultipleSelections: boolean;
  termSetId: string;
  anchorTermId?: string;
  panelTitle: string;
  label: string;
  context: BaseComponentContext;
  initialValues?: ITag[];
  errorMessage?: string; // TODO: is this needed?
  disabled?: boolean;
  required?: boolean;
}

export function ModernTaxonomyPicker(props: IModernTaxonomyPickerProps) {
  const [termsService] = React.useState(() => new SPTaxonomyService(props.context));
  const [termGroupId, setTermGroupId] = React.useState<string>();
  const [termSetName, setTermSetName] = React.useState<string>();

  const [terms, setTerms] = React.useState<ITermInfo[]>([]);
  const previousValues = React.useRef<IPickerTerms>([]);
  const [errorMessage, setErrorMessage] = React.useState(props.errorMessage);
  const [internalErrorMessage, setInternalErrorMessage] = React.useState<string>();
  const [panelIsOpen, setPanelIsOpen] = React.useState(false);
  const [resetOnClose, setResetOnClose] = React.useState(true); // was called cancel
  const [loading, setLoading] = React.useState(false); // was called loaded
  const [selectedOptions, setSelectedOptions] = React.useState<ITag[]>([]);
  const [selectedPanelOptions, setSelectedPanelOptions] = React.useState<ITag[]>([]);

  const invalidTerm = React.useRef<string>(null);

  React.useEffect(() => {
    sp.setup(props.context);
  }, []);

  // React.useEffect(() => {
  //   async function updateTermSetInfo(id: string): Promise<void> {
  //     if (!id) {
  //       return;
  //     }

  //     try {
  //       const termSetInfo: ITermSetInfo = await termsService.getTermSetInfo(id);
  //       // setTermSetId(termSetInfo.id);
  //       setTermGroupId(termSetInfo.groupId);
  //       setTermSetName(termSetInfo.localizedNames?.[0].name); // TODO: this should be changed to select a name based on localization
  //     } catch (error) {
  //     }
  //   }

  //   updateTermSetInfo(props.termSetId);
  // }, [props.termSetId]);

  React.useEffect(() => {
    setSelectedOptions(props.initialValues || []);
  }, [props.initialValues]);

  React.useEffect(() => {
    setErrorMessage(props.errorMessage);
  }, [props.errorMessage]);

  async function onOpenPanel(): Promise<void> {
    if (props.disabled === true) {
      return;
    }

    // Store the current code value
    // previousValues.current = cloneDeep(activeNodes);
    setResetOnClose(true);
    setLoading(true);

    const siteUrl = props.context.pageContext.site.absoluteUrl;
    const newTerms = await termsService.getTerms(Guid.parse(props.termSetId), Guid.empty, "", true, 50);
    setTerms(newTerms.value);

    setLoading(false);
    setPanelIsOpen(true);
  }

  function onClosePanel(): void {
    setLoading(false);
    setPanelIsOpen(false);
    // if (resetOnClose) {
    //   setActiveNodes(previousValues.current);
    // }
  }

  function validate(selectedTerms): void {
    // not yet implemented
  }

  function onSave(): void {
    // setResetOnClose(false);
    setSelectedOptions([...selectedPanelOptions]);
    onClosePanel();

    // validate(activeNodes);
  }

  function termsFromPickerChanged(pickerTerms: IPickerTerms) {
    // setActiveNodes(pickerTerms);
    validate(pickerTerms);
  }

  function validateInputText(): void {
    // Show error message, if any unresolved value exists inside taxonomy picker control
    if (!!invalidTerm.current) {
      // An unresolved value exists
      setErrorMessage(strings.TaxonomyPickerInvalidTerms.replace('{0}', invalidTerm.current));
    }
    else {
      // There are no unresolved values
      setErrorMessage(null);
    }
  }

  function onInputChange(input: string): string {
    // if (!input) {
    //   const { validateInput } = props;
    //   if (!!validateInput) {
    //     // Perform validation of input text, only if taxonomy picker is configured with validateInput={true} property.
    //     invalidTerm.current = null;
    //     validateInputText();
    //   }
    // }
    return input;
  }

  function onBlur(event: React.FocusEvent<HTMLElement | Autofill>): void {
    // const { validateInput } = props;
    // if (!!validateInput) {
    //   // Perform validation of input text, only if taxonomy picker is configured with validateInput={true} property.
    //   const target: HTMLInputElement = event.target as HTMLInputElement;
    //   const targetValue = !!target ? target.value : null;
    //   if (!!targetValue) {
    //     invalidTerm.current = targetValue;
    //   }
    //   else {
    //     invalidTerm.current = null;
    //   }
    //   validateInputText();
    // }
  }

  function termSetSelectedChange(termSet: ITermSet, isChecked: boolean): void {
    // not yet implemented
  }

  function termsChanged(term: ITerm, checked: boolean): void {
    // not yet implemented
  }

  async function updateTaxonomyTree(): Promise<void> {
    // not yet implemented
  }

  async function onResolveSuggestions(filter: string, selectedItems?: ITag[]): Promise<ITag[]> {
    const languageTag = props.context.pageContext.cultureInfo.currentUICultureName !== "" ? props.context.pageContext.cultureInfo.currentUICultureName : props.context.pageContext.web.languageName;
    if (filter === "") {
      return [];
    }
    const filteredTerms = await termsService.searchTerm(Guid.parse(props.termSetId), filter, languageTag, props.anchorTermId ? Guid.parse(props.anchorTermId) : undefined);
    const filteredTermsWithoutSelectedItems = filteredTerms.filter((term) => {
      if (!selectedItems || selectedItems.length === 0) {
        return true;
      }
      for (const selectedItem of selectedItems) {
        return selectedItem.key !== term.id;
      }
    });
    const filteredTermsAndAvailable = filteredTermsWithoutSelectedItems.filter((term) => term.isAvailableForTagging.filter((t) => t.setId === props.termSetId)[0].isAvailable);
    const filteredTags = filteredTermsAndAvailable.map((term) => {
      const key = term.id;
      const name = term.labels.filter((termLabel) => (languageTag === "" || termLabel.languageTag === languageTag) &&
        termLabel.name.toLowerCase().indexOf(filter.toLowerCase()) === 0)[0]?.name;
      return { key: key, name: name };
    });
    return filteredTags;
  }

  const {
    label,
    context,
    disabled,
    // isTermSetSelectable,
    allowMultipleSelections,
    // disabledTermIds,
    // disableChildrenOfDisabledParents,
    // placeholder,
    panelTitle,
    // anchorId,
    // termActions,
    required
  } = props;
  return (
    <div className={styles.modernTaxonomyPicker}>
      {label && <Label required={required}>{label}</Label>}
      <div className={styles.termField}>
        <div className={styles.termFieldInput}>
        <TagPicker

          removeButtonAriaLabel="Remove"
          onResolveSuggestions={onResolveSuggestions}
          itemLimit={allowMultipleSelections ? undefined : 1}
          selectedItems={selectedOptions}
          onChange={(itms?: ITag[]) => {
            setSelectedOptions(itms || []);
            setSelectedPanelOptions(itms || []);
          }}
          getTextFromItem={(tag: ITag, currentValue?: string) => tag.name}
          inputProps={{
            "aria-label": "Tag Picker",
            placeholder: "Ange en term som du vill tagga"
          }}
        />
        </div>
        <div className={styles.termFieldButton}>
          <IconButton disabled={disabled} iconProps={{ iconName: 'Tag' } as IIconProps} onClick={onOpenPanel} />
        </div>
      </div>

      <FieldErrorMessage errorMessage={errorMessage || internalErrorMessage} />

      <Panel
        // isOpen={openPanel}
        isOpen={panelIsOpen}
        hasCloseButton={true}
        onDismiss={onClosePanel}
        isLightDismiss={true}
        type={PanelType.medium}
        headerText={panelTitle}
        onRenderFooterContent={() => {
          const horizontalGapStackTokens: IStackTokens = {
            childrenGap: 10,
          };

          return (
            <Stack horizontal disableShrink tokens={horizontalGapStackTokens}>
              <PrimaryButton text={strings.SaveButtonLabel} value="Save" onClick={onSave} />
              <DefaultButton text={strings.CancelButtonLabel} value="Cancel" onClick={onClosePanel} />
            </Stack>
          );
        }}>

        {
          /* Show spinner in the panel while retrieving terms */
          loading === true ? <Spinner size={SpinnerSize.medium} /> : ''
        }
        {
          loading === false && props.termSetId && (
            <div key={props.termSetId} >
              <TaxonomyForm
                allowMultipleSelections={allowMultipleSelections}
                terms={terms}
                onResolveSuggestions={onResolveSuggestions}
                onLoadMoreData={termsService.getTerms}
                getTermSetInfo={termsService.getTermSetInfo}
                context={props.context}
                termSetId={Guid.parse(props.termSetId)}
                pageSize={50}
                selectedPanelOptions={selectedPanelOptions}
                setSelectedPanelOptions={setSelectedPanelOptions}
              />
            </div>
          )
        }
      </Panel>
    </div >
  );
}
