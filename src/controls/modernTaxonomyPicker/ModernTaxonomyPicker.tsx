import * as React from 'react';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { Guid } from '@microsoft/sp-core-library';
import { IIconProps } from 'office-ui-fabric-react/lib/components/Icon';
import { PrimaryButton, DefaultButton, IconButton, IButtonStyles } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IBasePickerStyleProps, IBasePickerStyles, ITag, TagPicker } from 'office-ui-fabric-react/lib/Pickers';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { IStyleFunctionOrObject } from 'office-ui-fabric-react/lib/Utilities';
import { sp } from '@pnp/sp';
import { SPTaxonomyService } from '../../services/SPTaxonomyService';
import { TaxonomyPanelContents } from './taxonomyPanelContents';
import styles from './ModernTaxonomyPicker.module.scss';
import * as strings from 'ControlStrings';
import { TooltipHost } from '@microsoft/office-ui-fabric-react-bundle';
import { useId } from '@uifabric/react-hooks';
import { ITooltipHostStyles } from 'office-ui-fabric-react';
import { ITermInfo, ITermSetInfo, ITermStoreInfo } from '@pnp/sp/taxonomy';

export interface IModernTaxonomyPickerProps {
  allowMultipleSelections: boolean;
  termSetId: string;
  anchorTermId?: string;
  panelTitle: string;
  label: string;
  context: BaseComponentContext;
  initialValues?: ITag[];
  disabled?: boolean;
  required?: boolean;
  onChange?: (newValue?: ITag[]) => void;
  placeHolder?: string;
}

export function ModernTaxonomyPicker(props: IModernTaxonomyPickerProps) {
  const [taxonomyService] = React.useState(() => new SPTaxonomyService(props.context));
  const [panelIsOpen, setPanelIsOpen] = React.useState(false);
  const [selectedOptions, setSelectedOptions] = React.useState<ITag[]>(Object.prototype.toString.call(props.initialValues) === '[object Array]' ? props.initialValues : []);
  const [selectedPanelOptions, setSelectedPanelOptions] = React.useState<ITag[]>([]);
  const [termStoreInfo, setTermStoreInfo] = React.useState<ITermStoreInfo>();
  const [termSetInfo, setTermSetInfo] = React.useState<ITermSetInfo>();
  const [anchorTermInfo, setAnchorTermInfo] = React.useState<ITermInfo>();

  React.useEffect(() => {
    sp.setup(props.context);
    taxonomyService.getTermStoreInfo()
      .then((localTermStoreInfo) => {
        setTermStoreInfo(localTermStoreInfo);
      });
    taxonomyService.getTermSetInfo(Guid.parse(props.termSetId))
      .then((localTermSetInfo) => {
        setTermSetInfo(localTermSetInfo);
      });
    if (props.anchorTermId && props.anchorTermId !== Guid.empty.toString()) {
      taxonomyService.getTermById(Guid.parse(props.termSetId), props.anchorTermId ? Guid.parse(props.anchorTermId) : Guid.empty)
      .then((localAnchorTermInfo) => {
        setAnchorTermInfo(localAnchorTermInfo);
      });
    }
  }, []);

  React.useEffect(() => {
    if (props.onChange) {
      props.onChange(selectedOptions);
    }
  }, [selectedOptions]);

  function onOpenPanel(): void {
    if (props.disabled === true) {
      return;
    }
    setSelectedPanelOptions(selectedOptions);
    setPanelIsOpen(true);
  }

  function onClosePanel(): void {
    setSelectedPanelOptions([]);
    setPanelIsOpen(false);
  }

  function onApply(): void {
    setSelectedOptions([...selectedPanelOptions]);
    onClosePanel();
  }

  async function onResolveSuggestions(filter: string, selectedItems?: ITag[]): Promise<ITag[]> {
    const languageTag = props.context.pageContext.cultureInfo.currentUICultureName !== '' ? props.context.pageContext.cultureInfo.currentUICultureName : termStoreInfo.defaultLanguageTag;
    if (filter === '') {
      return [];
    }
    const filteredTerms = await taxonomyService.searchTerm(Guid.parse(props.termSetId), filter, languageTag, props.anchorTermId ? Guid.parse(props.anchorTermId) : Guid.empty);
    const filteredTermsWithoutSelectedItems = filteredTerms.filter((term) => {
      if (!selectedItems || selectedItems.length === 0) {
        return true;
      }
      return selectedItems.every((item) => item.key !== term.id);
    });
    const filteredTermsAndAvailable = filteredTermsWithoutSelectedItems.filter((term) => term.isAvailableForTagging.filter((t) => t.setId === props.termSetId)[0].isAvailable);
    const filteredTags = filteredTermsAndAvailable.map((term) => {
      const key = term.id;
      let labelsWithMatchingLanguageTag = term.labels.filter((termLabel) => (termLabel.languageTag === languageTag));
      if (labelsWithMatchingLanguageTag.length === 0) {
        labelsWithMatchingLanguageTag = term.labels.filter((termLabel) => (termLabel.languageTag === termStoreInfo.defaultLanguageTag));
      }
      const name = labelsWithMatchingLanguageTag.filter((termLabel) => termLabel.name.toLowerCase().indexOf(filter.toLowerCase()) === 0)[0]?.name;
      return { key: key, name: name };
    });
    return filteredTags;
  }

  const calloutProps = { gapSpace: 0 };
  const tooltipId = useId('tooltip');
  const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
  const addTermButtonStyles: IButtonStyles = {rootHovered: {backgroundColor: "inherit"}, rootPressed: {backgroundColor: "inherit"}};
  const tagPickerStyles: IStyleFunctionOrObject<IBasePickerStyleProps, IBasePickerStyles> = { input: {minheight: 34}, text: {minheight: 34} };

  return (
    <div className={styles.modernTaxonomyPicker}>
      {props.label && <Label required={props.required}>{props.label}</Label>}
      <div className={styles.termField}>
        <div className={styles.termFieldInput}>
          <TagPicker
            removeButtonAriaLabel={strings.ModernTaxonomyPickerRemoveButtonText}
            onResolveSuggestions={onResolveSuggestions}
            itemLimit={props.allowMultipleSelections ? undefined : 1}
            selectedItems={selectedOptions}
            disabled={props.disabled}
            styles={tagPickerStyles}
            onChange={(itms?: ITag[]) => {
              setSelectedOptions(itms || []);
              setSelectedPanelOptions(itms || []);
            }}
            getTextFromItem={(tag: ITag) => tag.name}
            inputProps={{
              'aria-label': props.placeHolder || strings.ModernTaxonomyPickerDefaultPlaceHolder,
              placeholder: props.placeHolder || strings.ModernTaxonomyPickerDefaultPlaceHolder
            }}
          />
        </div>
        <div className={styles.termFieldButton}>
          <TooltipHost
            content={strings.ModernTaxonomyPickerAddTagButtonTooltip}
            id={tooltipId}
            calloutProps={calloutProps}
            styles={hostStyles}
          >
            <IconButton disabled={props.disabled} styles={addTermButtonStyles} iconProps={{ iconName: 'Tag' } as IIconProps} onClick={onOpenPanel} aria-describedby={tooltipId} />
          </TooltipHost>
        </div>
      </div>

      <Panel
        isOpen={panelIsOpen}
        hasCloseButton={true}
        closeButtonAriaLabel={strings.ModernTaxonomyPickerPanelCloseButtonText}
        onDismiss={onClosePanel}
        isLightDismiss={true}
        type={PanelType.medium}
        headerText={props.panelTitle}
        onRenderFooterContent={() => {
          const horizontalGapStackTokens: IStackTokens = {
            childrenGap: 10,
          };
          return (
            <Stack horizontal disableShrink tokens={horizontalGapStackTokens}>
              <PrimaryButton text={strings.ModernTaxonomyPickerApplyButtonText} value="Apply" onClick={onApply} />
              <DefaultButton text={strings.ModernTaxonomyPickerCancelButtonText} value="Cancel" onClick={onClosePanel} />
            </Stack>
          );
        }}>

        {
          props.termSetId && (
            <div key={props.termSetId} >
              <TaxonomyPanelContents
                allowMultipleSelections={props.allowMultipleSelections}
                onResolveSuggestions={onResolveSuggestions}
                onLoadMoreData={taxonomyService.getTerms}
                anchorTermInfo={anchorTermInfo}
                termSetInfo={termSetInfo}
                termStoreInfo={termStoreInfo}
                context={props.context}
                termSetId={Guid.parse(props.termSetId)}
                pageSize={50}
                selectedPanelOptions={selectedPanelOptions}
                setSelectedPanelOptions={setSelectedPanelOptions}
                placeHolder={props.placeHolder || strings.ModernTaxonomyPickerDefaultPlaceHolder}
              />
            </div>
          )
        }
      </Panel>
    </div >
  );
}