import * as React from 'react';
import styles from './TaxonomyForm.module.scss';
// import { ITaxonomyFormProps } from './ITaxonomyFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, ChoiceGroup, classNamesFunction, DetailsRow, GroupedList, GroupHeader, ICheckboxStyleProps, ICheckboxStyles, IChoiceGroupOption, IChoiceGroupOptionProps, IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles, IChoiceGroupStyleProps, IChoiceGroupStyles, IColumn, IGroup, IGroupedList, IGroupFooterProps, IGroupHeaderCheckboxProps, IGroupHeaderProps, IGroupRenderProps, IGroupShowAllProps, ILinkStyleProps, ILinkStyles, IRenderFunction, ISpinnerStyleProps, ISpinnerStyles, IStyleFunctionOrObject, ITag, Label, Link, Spinner, TagPicker } from 'office-ui-fabric-react';
// import { createListItems, createGroups, IExampleItem } from '@uifabric/example-data';
import { Selection, SelectionMode, SelectionZone } from 'office-ui-fabric-react/lib/Selection';
import { useBoolean, useConst } from '@uifabric/react-hooks';
import { ITermInfo, ITermSetInfo } from '@pnp/sp/taxonomy';
import { Guid } from '@microsoft/sp-core-library';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface ITaxonomyFormProps {
  // termSetName: string;
  // termSetId: string;
  context: BaseComponentContext;
  multiSelection: boolean;
  terms: ITermInfo[];
  termSetId: Guid;
  pageSize: number;
  onResolveSuggestions: (filter: string, selectedItems?: ITag[]) => ITag[] | PromiseLike<ITag[]>;
  onLoadMoreData: (termSetId: Guid, parentTermId?: Guid, skiptoken?: string, hideDeprecatedTerms?: boolean, pageSize?: number) => Promise<{ value: ITermInfo[], skiptoken: string }>;
  getTermSetInfo: (termSetId: Guid) => Promise<ITermSetInfo | undefined>;
}

// const multiSelect = true;
// const groupCount = 3;
// const groupDepth = 3;
// const items: IExampleItem[] = createListItems(Math.pow(groupCount, groupDepth + 1));
// const groups = createGroups(groupCount, groupDepth, 0, groupCount, 1, "Test", true);
// const columns = Object.keys(items[0])
//   .slice(0, 3)
//   .map(
//     (key: string): IColumn => ({
//       key: key,
//       name: key,
//       fieldName: key,
//       minWidth: 300,
//     }),
//   );
// let mockGroups: IGroup[] = [];

export function TaxonomyForm(props: ITaxonomyFormProps): React.ReactElement<ITaxonomyFormProps> {
  const groupedListRef = React.useRef<IGroupedList>();

  const [selectedOptions, setSelectedOptions] = React.useState<ITag[]>([]);
  const [groupsLoading, setGroupsLoading] = React.useState<string[]>([]);
  const [groups, setGroups] = React.useState<IGroup[]>([]);


  React.useEffect(() => {
    // if (props.terms != null) {
    // let group: IGroup = { name: name, key: name, startIndex: -1, count: 50, level: 1, isCollapsed: true, data: {}, hasMoreData: true };
    setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, props.termSetId.toString()]);

    props.getTermSetInfo(props.termSetId)
      .then((termSetInfo) => {
        const languageTag = props.context.pageContext.cultureInfo.currentUICultureName !== "" ? props.context.pageContext.cultureInfo.currentUICultureName : props.context.pageContext.web.languageName;

        const termSetName = termSetInfo.localizedNames.filter((name) => name.languageTag === languageTag)[0].name;
        const rootGroup: IGroup = { name: termSetName, key: termSetInfo.id, startIndex: -1, count: 50, level: 0, isCollapsed: false, data: {skiptoken: ""}, hasMoreData: false };
        setGroups([rootGroup]);
        props.onLoadMoreData(props.termSetId, Guid.empty, "", true)
        .then((terms) => {
          const grps: IGroup[] = terms.value.map(term => {
            const g: IGroup = {
              name: term.labels?.[0].name, // TODO: fix this by looking up correct language
              key: term.id,
              startIndex: -1,
              count: 50,
              level: 1,
              isCollapsed: true,
              data: {skiptoken: "", term: term},
              hasMoreData: term.childrenCount > 0,
            };
            if (g.hasMoreData) {
              g.children = [];
            }
            return g;
          });
          rootGroup.children = grps;
          rootGroup.data.skiptoken = terms.skiptoken;
          rootGroup.hasMoreData = terms.skiptoken !== "";
          setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== props.termSetId.toString()));
          setGroups([rootGroup]);
        });

      });
  }, []);

  const onToggleCollapse = (group: IGroup): void => {
    if (group.isCollapsed === true) {
      setGroups((prevGroups) => {
        const recurseGroups = (currentGroup) => {
          if (currentGroup.key === group.key) {
            currentGroup.isCollapsed = false;
          }
          if (currentGroup.children?.length > 0) {
            for (const child of currentGroup.children) {
              recurseGroups(child);
            }
          }
        }
        let newGroupsState: IGroup[] = [];
        for (const prevGroup of prevGroups) {
          recurseGroups(prevGroup);
          newGroupsState.push(prevGroup);
        }

        return newGroupsState;
      });

      if (group.children && group.children.length === 0) {
        setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, group.key]);
        group.data.isLoading = true;

        props.onLoadMoreData(props.termSetId, Guid.parse(group.key), "", true)
          .then((terms) => {
            const grps: IGroup[] = terms.value.map(term => {
              const g: IGroup = {
                name: term.labels?.[0].name, // TODO: fix this by looking up correct language
                key: term.id,
                startIndex: -1,
                count: 50,
                level: group.level + 1,
                isCollapsed: true,
                data: {skiptoken: "", term: term},
                hasMoreData: term.childrenCount > 0,
              };
              if (g.hasMoreData) {
                g.children = [];
              }
              return g;
            });
            group.children = grps;
            group.data.skiptoken = terms.skiptoken;
            group.hasMoreData = terms.skiptoken !== "";
            setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== group.key));
        });
      }
    }
    else {
      setGroups((prevGroups) => {
        const recurseGroups = (currentGroup) => {
          if (currentGroup.key === group.key) {
            currentGroup.isCollapsed = true;
          }
          if (currentGroup.children?.length > 0) {
            for (const child of currentGroup.children) {
              recurseGroups(child);
            }
          }
        }
        let newGroupsState: IGroup[] = [];
        for (const prevGroup of prevGroups) {
          recurseGroups(prevGroup);
          newGroupsState.push(prevGroup);
        }

        return newGroupsState;
      });

    }
  };

  const onChoiceChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void => {
    setSelectedOptions([{ key: option.key, name: option.text }]);
  };

  const onCheckboxChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean, tag?: ITag): void => {
    if (checked) {
      setSelectedOptions((prevOptions) => [...prevOptions, tag]);
    }
    else {
      setSelectedOptions((prevOptions) => prevOptions.filter((value) => value.key !== tag.key));
    }
  };

  const onRenderTitle = (groupHeaderProps: IGroupHeaderProps) => {
    if (groupHeaderProps.group.level === 0) {
      return (
        <Label>{groupHeaderProps.group.name}</Label>
      );
    }
    if (props.multiSelection) {
      const isSelected = selectedOptions.some(value => value.key === groupHeaderProps.group.key);
      const selectedStyles: IStyleFunctionOrObject<ICheckboxStyleProps, ICheckboxStyles> = isSelected ? { label: { fontWeight: "bold", color: "#000000" } } : { label: { fontWeight: "normal", color: "#000000" } };
      return (
        <Checkbox
          key={groupHeaderProps.group.key}
          label={groupHeaderProps.group.name}
          onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) =>
            onCheckboxChange(ev, checked, { name: groupHeaderProps.group.name, key: groupHeaderProps.group.key })}
          checked={isSelected}
          styles={selectedStyles}
          disabled={groupHeaderProps.group.data.term.isAvailableForTagging.filter((t) => t.setId === props.termSetId.toString())[0].isAvailable === false}
        />
      );
    }
    else {
      const isSelected = selectedOptions?.[0]?.key === groupHeaderProps.group.key;
      const selectedStyle: IStyleFunctionOrObject<IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles> = isSelected ? { choiceFieldWrapper: { fontWeight: "bold" }, labelWrapper: {color: "#000000"} } : { choiceFieldWrapper: { fontWeight: "normal" }, labelWrapper: {color: "#000000"} };
      const getClassNames = classNamesFunction<IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles>();

      const LARGE_IMAGE_SIZE = 71;

      const DEFAULT_PROPS: Partial<IChoiceGroupOptionProps> = {
        // This ensures default imageSize value doesn't mutate. Mutation can cause style re-calcuation.
        imageSize: { width: 32, height: 32 },
      };

      const classNames = getClassNames(selectedStyle!, {
        theme: undefined,
        hasIcon: true,
        hasImage: true,
        checked: false,
        disabled: false,
        imageIsLarge: false,
        imageSize: undefined,
        focused: false,
      });

      const options: IChoiceGroupOption[] = [{ key: groupHeaderProps.group.key, text: groupHeaderProps.group.name, styles: selectedStyle, onRenderLabel: (p) => <label htmlFor={p.id} className={classNames.field}><span id={p.labelId} /*style={}*/ color={"#000000"}>{p.text}</span></label> }];
      return (
        <ChoiceGroup
          options={options}
          selectedKey={selectedOptions?.[0]?.key}
          onChange={onChoiceChange}
          disabled={groupHeaderProps.group.data.term.isAvailableForTagging.filter((t) => t.setId === props.termSetId.toString())[0].isAvailable === false}
        />
      );
    }
  };

  const onRenderHeader = (headerProps: IGroupHeaderProps): JSX.Element => {
    const headerCountStyle = { "display": "none" };
    const checkButtonStyle = { "display": "none" };
    const expandStyle = { "visibility": "hidden" };

    return (
      <GroupHeader
        {...headerProps}
        styles={{
          "expand": !headerProps.group.children || headerProps.group.level === 0 ? expandStyle : null,
          "expandIsCollapsed": !headerProps.group.children || headerProps.group.level === 0 ? expandStyle : null,
          "check": checkButtonStyle,
          "headerCount": headerCountStyle,
        }}
        onRenderTitle={onRenderTitle}
        onToggleCollapse={onToggleCollapse}
        indentWidth={20}
      />
    );
  };

  const onRenderFooter = (footerProps: IGroupFooterProps): JSX.Element => {
    if ((footerProps.group.hasMoreData || footerProps.group.children && footerProps.group.children.length === 0) && !footerProps.group.isCollapsed) {
      if (groupsLoading.some(value => value === footerProps.group.key)) {
        const spinnerStyles: IStyleFunctionOrObject<ISpinnerStyleProps, ISpinnerStyles> = { circle: { verticalAlign: "middle" } };
        return (
          <div style={{ height: "48px", lineHeight: "48px", display: "flex", justifyContent: "center", alignItems: "center" }}>
            <Spinner styles={spinnerStyles} />
          </div>
        );
      }
      const linkStyles: IStyleFunctionOrObject<ILinkStyleProps, ILinkStyles> = { root: { fontSize: "14px", paddingLeft: (footerProps.groupLevel + 1) * 20 + 62 } };
      return (
        <div style={{ height: "48px", lineHeight: "48px" }}>
          <Link onClick={() => {
              props.onLoadMoreData(props.termSetId, footerProps.group.key === props.termSetId.toString() ? Guid.empty : Guid.parse(footerProps.group.key), footerProps.group.data.skiptoken, true)
              .then((terms) => {
                const grps: IGroup[] = terms.value.map(term => {
                  const g: IGroup = {
                    name: term.labels?.[0].name, // TODO: fix this by looking up correct language
                    key: term.id,
                    startIndex: -1,
                    count: 50,
                    level: footerProps.group.level + 1,
                    isCollapsed: true,
                    data: {skiptoken: "", term: term},
                    hasMoreData: term.childrenCount > 0,
                  };
                  if (g.hasMoreData) {
                    g.children = [];
                  }
                  return g;
                });
                footerProps.group.children = [...footerProps.group.children, ...grps];
                footerProps.group.data.skiptoken = terms.skiptoken;
                footerProps.group.hasMoreData = terms.skiptoken !== "";
                setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== footerProps.group.key));
              });
            }}
            styles={linkStyles}>
            Load more...
          </Link>
        </div>
      );
    }
    return null;
  };

  const onRenderShowAll: IRenderFunction<IGroupShowAllProps> = () => {
    return null;
  };

  const groupProps: IGroupRenderProps = {
    onRenderFooter: onRenderFooter,
    onRenderHeader: onRenderHeader,
    showEmptyGroups: true,
    onRenderShowAll: onRenderShowAll,
  };

  function getTagText(tag: ITag, currentValue?: string) {
    return tag.name;
  }

  const onPickerChange = (itms?: ITag[]): void => {
    // const newSelection = itms.map(value => value.key.toString());
    setSelectedOptions(itms || []);
  };

  return (
    <div className={styles.taxonomyForm}>
      {/* <div className={styles.container}> */}
        {/* <div className={styles.row}> */}
          {/* <div className={styles.column}> */}
            {/* <span className={styles.title}>Welcome to SharePoint!</span>
            <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
            <p className={styles.description}>{escape(props.description)}</p> */}
            <TagPicker
              removeButtonAriaLabel="Remove"
              onResolveSuggestions={props.onResolveSuggestions}
              itemLimit={props.multiSelection ? undefined : 1}
              // selectedItems={getITagsFromSelected()}
              selectedItems={selectedOptions}
              onChange={onPickerChange}
              getTextFromItem={getTagText}
            />
            <GroupedList
              componentRef={groupedListRef}
              items={[]}
              onRenderCell={null}
              groups={groups}
              groupProps={groupProps}
            />
            {/* {(selectedOptions.map((value) => { return <><p className={styles.description}>{value}</p><br /></>; }))} */}
          {/* </div> */}
        {/* </div> */}
      {/* </div> */}
    </div>
  );
}

