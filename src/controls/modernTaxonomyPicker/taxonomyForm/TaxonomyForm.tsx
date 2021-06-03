import * as React from 'react';
import styles from './TaxonomyForm.module.scss';
// import { ITaxonomyFormProps } from './ITaxonomyFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, ChoiceGroup, DetailsRow, GroupedList, GroupHeader, ICheckboxStyleProps, ICheckboxStyles, IChoiceGroupOption, IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles, IChoiceGroupStyleProps, IChoiceGroupStyles, IColumn, IGroup, IGroupedList, IGroupFooterProps, IGroupHeaderCheckboxProps, IGroupHeaderProps, IGroupRenderProps, IGroupShowAllProps, ILinkStyleProps, ILinkStyles, IRenderFunction, ISpinnerStyleProps, ISpinnerStyles, IStyleFunctionOrObject, ITag, Label, Link, Spinner, TagPicker } from 'office-ui-fabric-react';
// import { createListItems, createGroups, IExampleItem } from '@uifabric/example-data';
import { Selection, SelectionMode, SelectionZone } from 'office-ui-fabric-react/lib/Selection';
import { useBoolean, useConst } from '@uifabric/react-hooks';
import { ITermInfo } from '@pnp/sp/taxonomy';

export interface ITaxonomyFormProps {
  // termSetName: string;
  // termSetId: string;
  multiSelection: boolean;
  terms: ITermInfo[];
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
    const grps: IGroup[] = (props.terms || []).map(term => {
      const g: IGroup = {
        name: term.labels?.[0].name, // TODO: fix this by looking up correct language
        key: term.id,
        startIndex: -1,
        count: 50,
        level: 1,
        isCollapsed: true,
        data: {},
        hasMoreData: term.childrenCount > 0,
      };
      if (g.hasMoreData) {
        g.children = [];
      }
      return g;
    });
    const rootGroup: IGroup = { name: "TermSet name", key: "TermSet name", startIndex: -1, count: 50, level: 0, isCollapsed: false, data: {}, hasMoreData: true };
    rootGroup.children = grps;
    // setGroups(grps);
    setGroups([rootGroup]);

    // } else {
    //   setGroups([]);
    // }

  }, [props.terms]);

  // TODO: ta bort när laddning är på plats
  const generateChildGroups = (group: IGroup, level: number): IGroup[] => {
    let childGroups: IGroup[] = [];
    for (let i = 0; i < 3; i++) {
      const name = `${group.name} - ${i}`;
      let newGroup: IGroup = { name: name, key: name, startIndex: -1, count: 50, level: level, isCollapsed: true, hasMoreData: i === 1 ? true : false, data: {} };
      if (level <= 3) {
        newGroup.children = [];
      }
      childGroups.push(newGroup);
    }
    return childGroups;
  };


  const onToggleCollapse = (group: IGroup): void => {
    if (group.isCollapsed === true) {
      group.isCollapsed = false;

      if (group.children && group.children.length === 0) {
        setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, group.key]);
        group.data.isLoading = true;

        // TODO: ladda data för term
        setTimeout((loadedGroup) => {
          group.children = generateChildGroups(loadedGroup, loadedGroup.level + 1); // själva laddningen
          // efter laddning av data
          setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== loadedGroup.key));
          loadedGroup.data.nextPage = ""; // odata.nextlink
          loadedGroup.data.isLoading = false;
          groupedListRef.current.forceUpdate();
        }, 3000, group);
      }
      groupedListRef.current.forceUpdate();
    }
    else {
      group.isCollapsed = true;
    }
    groupedListRef.current.forceUpdate();
  };

  const onChoiceChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void => {
    setSelectedOptions([{ key: option.key, name: option.text }]);
    groupedListRef.current.forceUpdate();
  };

  const onCheckboxChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean, tag?: ITag): void => {
    if (checked) {
      setSelectedOptions((prevOptions) => [...prevOptions, tag]);
    }
    else {
      setSelectedOptions((prevOptions) => prevOptions.filter((value) => value.key !== tag.key));
    }
    groupedListRef.current.forceUpdate();
  };

  const onRenderTitle = (groupHeaderProps: IGroupHeaderProps) => {
    if (groupHeaderProps.group.level === 0) {
      return (
        <Label>{groupHeaderProps.group.name}</Label>
      );
    }
    if (props.multiSelection) {
      const isSelected = selectedOptions.some(value => value.key === groupHeaderProps.group.key);
      const selectedStyles: IStyleFunctionOrObject<ICheckboxStyleProps, ICheckboxStyles> = isSelected ? { label: { fontWeight: "bold" } } : { label: { fontWeight: "normal" } };
      return (
        <Checkbox
          id={groupHeaderProps.group.key}
          key={groupHeaderProps.group.key}
          label={groupHeaderProps.group.name}
          onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) =>
            onCheckboxChange(ev, checked, { name: groupHeaderProps.group.name, key: groupHeaderProps.group.key })}
          checked={isSelected}
          styles={selectedStyles}
        />
      );
    }
    else {
      const isSelected = selectedOptions?.[0]?.key === groupHeaderProps.group.key;
      const selectedStyle: IStyleFunctionOrObject<IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles> = isSelected ? { choiceFieldWrapper: { fontWeight: "bold" } } : { choiceFieldWrapper: { fontWeight: "normal" } };

      const options: IChoiceGroupOption[] = [{ key: groupHeaderProps.group.key, text: groupHeaderProps.group.name, styles: selectedStyle }];
      return (
        <ChoiceGroup options={options} selectedKey={selectedOptions?.[0]?.key} onChange={onChoiceChange} />
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

  const onLoadMoreClick = (group: IGroup) => {
    setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, group.key]);
    group.data.isLoading = true;

    setTimeout((loadedGroup) => {
      loadedGroup.children = [...loadedGroup.children, ...generateChildGroups(loadedGroup, loadedGroup.level + 1)];
      setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== loadedGroup.key));
      loadedGroup.data.nextPage = "";
      loadedGroup.data.isLoading = false;
      groupedListRef.current.forceUpdate();
    }, 3000, group);
    groupedListRef.current.forceUpdate();
  };

  const onRenderFooter = (footerProps: IGroupFooterProps): JSX.Element => {
    if (footerProps.group.data) {

    }
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
          <Link onClick={() => onLoadMoreClick(footerProps.group)} styles={linkStyles}>
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

  // // Denna useEffect används bara för att fixa till det mockade datat
  // React.useEffect(() => {
  //   let rootGroup: IGroup = { name: "TermSet name", key: "TermSet name", startIndex: -1, count: 50, level: 0, isCollapsed: false, data: {}, hasMoreData: true };
  //   rootGroup.children = [];

  //   for (let i = 0; i < 3; i++) {
  //     const name = `Testar - 0 - ${i}`;
  //     let group: IGroup = { name: name, key: name, startIndex: -1, count: 50, level: 1, isCollapsed: true, data: {}, hasMoreData: true };
  //     group.children = [];
  //     rootGroup.children.push(group);
  //   }
  //   mockGroups.push(rootGroup);
  //   groupedListRef.current.forceUpdate();
  // }, []);

  // const getITagsFromSelected = (): ITag[] => {
  //   let iTags: ITag[] = [];
  //   for (const selectedOption of selectedOptions) {
  //     iTags.push({ key: selectedOption, name: selectedOption });
  //   }
  //   return iTags;
  // };

  const onResolveSuggestions = (filter: string, selectedItems?: ITag[]): ITag[] | PromiseLike<ITag[]> => {
    return groups[0].children.filter(group => group.name.toLowerCase().indexOf(filter.toLowerCase()) === 0).map(g => ({ key: g.key, name: g.name }));
  };

  const onPickerChange = (itms?: ITag[]): void => {
    // const newSelection = itms.map(value => value.key.toString());
    setSelectedOptions(itms || []);
    groupedListRef.current.forceUpdate();
  };

  return (
    <div className={styles.taxonomyForm}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            {/* <span className={styles.title}>Welcome to SharePoint!</span>
            <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
            <p className={styles.description}>{escape(props.description)}</p> */}
            <TagPicker
              removeButtonAriaLabel="Remove"
              onResolveSuggestions={onResolveSuggestions}
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
          </div>
        </div>
      </div>
    </div>
  );
}

