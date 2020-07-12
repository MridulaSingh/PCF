import * as React from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react/lib/ScrollablePane';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, IDetailsFooterProps, Selection, SelectionMode, IDetailsHeaderProps, IGroup, IGroupDividerProps, DetailsHeader } from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyleSets, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

//#region Style Constants

const classNames = mergeStyleSets({
    item: {
        selectors: {
            '&:hover': {
                cursor: 'pointer'
            }
        }
    },
    listFooter: {
        display: 'flex',
        padding: '1px'
    },
    cmdBarFarItems: {
        fontSize: '0.857143rem',

    },
    labelClass: {
        fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;',
        fontSize: '15px', padding: '2px', paddingTop: '15px', paddingLeft: '30px'
    }
});
const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 200 },
    //label: { padding: 'inherits' },
    // panel: { padding: 'inherits' },
    root: { display: 'flex', padding: '10px' }
};

const controlWrapperClass = mergeStyles({
    display: 'flex',
    flexWrap: 'wrap',
    //display: 'inline-table'
});
//#endregion

//#region Interfaces
export interface IGroupDetailListControlProps {
    data: IListData[];
    columns: IListColumn[];
    groupBy?: IGroup[];
    groupColumn?: any;
    totalResultCount: number;
    allocatedWidth: number;
    triggerNavigate?: (id: string) => void;
    triggerSelection?: (selectedKeys: any[]) => void;
}

export interface IListData {
    attribute: string;
    value: string;
}

export interface IListColumn extends IColumn {
    dataType?: string,
    isPrimary?: boolean
}

export interface IListControlState extends React.ComponentState {
}
//#endregion 

export class GroupListControl extends React.Component<IGroupDetailListControlProps, IListControlState> {

    //#region Global Variables
    private _selection: Selection;
    private _totalWidth: number;
    private _cmdBarFarItems: ICommandBarItemProps[];
    private _cmdBarItems: ICommandBarItemProps[];
    private _totalRecords: number;
    private _groupingColumn?: string = "No-Grouping";
    //#endregion

    constructor(props: IGroupDetailListControlProps) {
        super(props);

        this._totalWidth = this._totalColumnWidth(props.columns);
        this._totalWidth = this._totalWidth > props.allocatedWidth ? this._totalWidth : props.allocatedWidth;
        this._totalRecords = props.totalResultCount;

        this.state = {
            _items: props.data,
            _columns: this._buildColumns(props.columns),
            _groupBy: this._groupBy(props.data, this._groupingColumn),
            _group: this._groupingColumn,
            _triggerNavigate: props.triggerNavigate,
            _triggerSelection: props.triggerSelection,
            _selectionCount: 0
        };

        this._selection = new Selection({
            onSelectionChanged: () => {
                this.setState({
                    _selectionCount: this._setSelectionDetails(),
                });
            }
        });
        this._cmdBarFarItems = this.renderCommandBarFarItem(props.data.length);
        this._cmdBarItems = [];
    }


    public componentWillReceiveProps(newProps: IGroupDetailListControlProps): void {
        this.setState({
            _items: newProps.data,
            _columns: this._buildColumns(newProps.columns),
            _groupBy: this._groupBy(newProps.data, this._groupingColumn)
        });
        this._totalWidth = this._totalColumnWidth(newProps.columns);
        this._totalRecords = newProps.totalResultCount;
        this._cmdBarFarItems = this.renderCommandBarFarItem(newProps.data.length);
    }

    private _dropdownOptions(): IDropdownOption[] {
        var columns = this.props.columns;
        let iDropdowns: IDropdownOption[] = [];
        for (var column of columns) {
            let iColumn: IDropdownOption = {
                key: column.key,
                text: column.name
            }
            iDropdowns.push(iColumn);
        }
        var defaultdropdown: IDropdownOption = { key: "No-Grouping", text: "No-Grouping", selected: true };
        iDropdowns.push(defaultdropdown);
        return iDropdowns;

    }

    private _onRenderDetailsHeader = (props: IDetailsHeaderProps | undefined, defaultRender?: IRenderFunction<IDetailsHeaderProps>): JSX.Element => {
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={false}>
                <div className={controlWrapperClass} style={{ display: "flex" }}>
                    <label className={classNames.labelClass}>Group By: </label>
                    <Dropdown
                        //label="Group By : "
                        defaultSelectedKey={this._groupingColumn}
                        options={this._dropdownOptions()}
                        styles={dropdownStyles}
                        onChange={(e, selectedOption: any) => {
                            this._groupingColumn = selectedOption?.key;
                            this.setState({ _groupBy: this._groupBy(this.state._items, selectedOption?.key) });
                        }}
                    />
                </div>
                <DetailsHeader {...props} ariaLabelForToggleAllGroupsButton={'Expand collapse groups'} layoutMode={DetailsListLayoutMode.justified} />
            </Sticky>
        );
    }


    private _onRenderDetailsFooter = (props: IDetailsFooterProps | undefined, defaultRender?: IRenderFunction<IDetailsFooterProps>): JSX.Element => {
        return (
            <Sticky stickyPosition={StickyPositionType.Footer} isScrollSynced={true}>
                <div className={classNames.listFooter}>
                    <Label style={{ padding: 'inherits', margin: 'inherits' }} className={"listFooterLabel"}>{`${this.state._selectionCount} selected`}</Label>
                    <CommandBar className={"cmdbar"} farItems={this._cmdBarFarItems} items={this._cmdBarItems} />
                </div>
            </Sticky>
        );
    }

    private _onColumnClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn): void => {
        let updatedColumns: IColumn[] = this.state._columns;
        let sortedItems: IListData[] = this.state._items;
        let isSortedDescending: boolean | undefined = column?.isSortedDescending;
        this._groupingColumn = (this._groupingColumn == "No-Grouping") ? "No-Grouping" : column?.key
        if (column?.isSorted) {
            isSortedDescending = !isSortedDescending;
        }

        sortedItems = this._sort(sortedItems, column?.fieldName!, isSortedDescending);

        this.setState({
            _items: sortedItems,
            _columns:
                updatedColumns.map(col => {
                    col.isSorted = col.key === column?.key;
                    if (col.isSorted) {
                        col.isSortedDescending = isSortedDescending;
                    }
                    return col;
                }),
            _groupBy: this._groupBy(sortedItems, this._groupingColumn, isSortedDescending)
        });
        
    }

    private _setSelectionDetails(): number {
        let selectedKeys = [];
        let selections = this._selection.getSelection();
        for (let selection of selections) {
            selectedKeys.push(selection.key as string);
        }

        this.state._triggerSelection(selectedKeys);

        switch (selectedKeys.length) {
            case 0:
                return 0;
            default:
                return selectedKeys.length;
        }
    }

    private _sort = <T,>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] => {
        let key = columnKey as keyof T;
        return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    }

    private renderCommandBarFarItem(recordsLoaded: number): ICommandBarItemProps[] {
        return [
            {
                key: 'total',
                text: `Total Records: ${this._totalRecords}`,
                ariaLabel: 'Total',
                iconProps: { iconName: 'ChevronRight' },
                disabled: recordsLoaded == this._totalRecords,
                className: classNames.cmdBarFarItems
            }
        ];
    }

    private _onItemInvoked(item: any): void {
        this.state._triggerNavigate(item.key);
    }

    private _buildColumns(listData: IListColumn[]): IColumn[] {
        let iColumns: IColumn[] = [];

        for (var column of listData) {
            let iColumn: IColumn = {
                key: column.key,
                name: column.name,
                fieldName: column.fieldName,
                currentWidth: column.currentWidth,
                minWidth: column.minWidth,
                maxWidth: column.maxWidth,
                isResizable: column.isResizable,
                sortAscendingAriaLabel: column.sortAscendingAriaLabel,
                sortDescendingAriaLabel: column.sortDescendingAriaLabel,
                className: column.className,
                headerClassName: column.headerClassName,
                data: column.data,
                isSorted: column.isSorted,
                isSortedDescending: column.isSortedDescending,

            }

            //create links for primary field and entity reference.            
            if (column.dataType && (column.dataType === "Lookup" || column.isPrimary)) {
                iColumn.onRender = (item: any, index: number | undefined, column: IColumn | undefined) => (
                    <Link key={item.key} onClick={() => this.state._triggerNavigate(item.key)}>{item[column?.fieldName!]}</Link>
                );
            }
            else if (column.dataType === "Email") {
                iColumn.onRender = (item: any, index: number | undefined, column: IColumn | undefined) => (
                    <Link href={`mailto:${item[column?.fieldName!]}`} >{item[column?.fieldName!]}</Link>
                );
            }
            else if (column.dataType === "Phone") {
                iColumn.onRender = (item: any, index: number | undefined, column: IColumn | undefined) => (
                    <Link href={`skype:${item[column?.fieldName!]}?call`} >{item[column!.fieldName!]}</Link>
                );
            }

            iColumns.push(iColumn);
        }

        return iColumns;
    }

    private _groupBy(items: any, fieldName?: string, isSortedDesending?: boolean) {
        let groups: any;

        if (fieldName == undefined || fieldName == "No-Grouping")
            groups = null;
        else {
            let data: any = this._sort(items, fieldName, isSortedDesending);
            groups = data.reduce((currentGroups: any, currentItem: { [x: string]: any; }, index: any) => {
                let lastGroup = currentGroups[currentGroups.length - 1];
                let fieldValue = currentItem[fieldName];

                if (!lastGroup || lastGroup.value !== fieldValue) {
                    currentGroups.push({
                        key: 'group' + fieldValue + index,
                        name: `By "${fieldValue}"`,
                        value: fieldValue,
                        startIndex: index,
                        level: 0,
                        count: 0,
                        isCollapsed: true,
                    });
                }
                if (lastGroup) {
                    lastGroup.count = index - lastGroup.startIndex;
                }
                return currentGroups;
            }, []);

            // Fix last group count
            let lastGroup = groups[groups.length - 1];

            if (lastGroup) {
                lastGroup.count = data.length - lastGroup.startIndex;
            }
        }
        return groups;
    }

    private _totalColumnWidth(listData: IListColumn[]): number {
        let totalColumnWidth: number;

        totalColumnWidth = listData
            .map(v => v.maxWidth!)
            .reduce((sum, current) => sum + current);

        // Add extra buffer
        return totalColumnWidth + 100;
    }

    //#endregion

    //#region Main Render Function
    public render() {
        return (
            <Fabric>

                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                    <DetailsList
                        setKey="parentcustomerid"
                        items={this.state._items}
                        columns={this.state._columns}
                        groups={this.state._groupBy}
                        onColumnHeaderClick={this._onColumnClick}
                        layoutMode={DetailsListLayoutMode.justified}
                        constrainMode={ConstrainMode.unconstrained}
                        onItemInvoked={this._onItemInvoked}
                        selection={this._selection}
                        selectionPreservedOnEmptyClick={true}
                        selectionMode={SelectionMode.multiple}
                        onRenderDetailsHeader={this._onRenderDetailsHeader}
                        onRenderDetailsFooter={this._onRenderDetailsFooter}
                        ariaLabelForSelectionColumn="Toggle selection"
                        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                        checkButtonAriaLabel="Row checkbox"
                    />
                </ScrollablePane>
            </Fabric>
        );
    }
    //#endregion
}
