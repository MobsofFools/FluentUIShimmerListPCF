//#region Imports

import * as React from "react";
import { useState, useEffect, useRef, useCallback } from "react";
import { useBoolean } from "@fluentui/react-hooks";

import { IRenderFunction } from "@fluentui/react/lib/Utilities";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import {
  ScrollablePane,
  ScrollbarVisibility,
} from "@fluentui/react/lib/ScrollablePane";
import {
  IDetailsHeaderProps,
  IDetailsColumnRenderTooltipProps,
  Selection,
  SelectionMode,
  IColumn,
  IDetailsRowProps,
} from "@fluentui/react/lib/DetailsList";

import { ShimmeredDetailsList } from "@fluentui/react/lib/ShimmeredDetailsList";
import { Sticky, StickyPositionType } from "@fluentui/react/lib/Sticky";
import { Panel } from "@fluentui/react/lib/Panel";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { Stack, StackItem } from "@fluentui/react/lib/Stack";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MarqueeSelection } from "@fluentui/react/lib/MarqueeSelection";
import {
  CommandBar,
  ICommandBarItemProps,
} from "@fluentui/react/lib/CommandBar";

//#endregion

//#region Props
type Props = {
  _context: ComponentFramework.WebApi;
};
//#endregion

//#region Original Header
const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (
  props,
  defaultRender
) => {
  if (!props) {
    return null;
  }
  const onRenderColumnHeaderTooltip: IRenderFunction<IDetailsColumnRenderTooltipProps> =
    (tooltipHostProps) => <TooltipHost {...tooltipHostProps} />;
  return (
    <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
      {defaultRender!({
        ...props,
        onRenderColumnHeaderTooltip,
      })}
    </Sticky>
  );
};
//#endregion

const ShimmerListComponent: React.FC<Props> = ({ _context }) => {
  //#region Column Resize

  const getWidth = () => modalRef.current?.clientWidth;
  const modalRef = useRef<HTMLDivElement>(null);
  const [modalWidth, setModalWidth] = useState(0);
  useEffect(() => {
    const resizeListener = () => {
      let width = getWidth();
      if (width) {
        setModalWidth(width);
      }
    };
    window.addEventListener("resize", resizeListener);
    return () => {
      // remove resize listener
      window.removeEventListener("resize", resizeListener);
    };
  }, []);
  useEffect(() => {
    if (modalRef.current) {
      let width = modalRef.current.clientWidth;
      setModalWidth(width);
    }
  }, [modalRef]);

  //#endregion

  //#region Column Sorting
  function unsortAll(column:string) {
    switch(column){
      case "name":
        setAddressSorted(false);
        setCitySorted(false);
        setCountrySorted(false);
        break;
      case "address":
        setNameSorted(false);
        setCitySorted(false);
        setCountrySorted(false);
        break;
      case "city":
        setNameSorted(false);
        setAddressSorted(false);
        setCountrySorted(false);
        break;
      case "country":
        setNameSorted(false);
        setAddressSorted(false);
        setCitySorted(false);
        break;
    }
  }

  const [isNameSorted, setNameSorted] = useState(false);
  const [isNameSortedDesc, setNameSortedDesc] = useState(false);
  const onNameColumnClick = () => {
    isNameSorted?
      isNameSortedDesc? 
      ascendingSortedCall("name")
      .then(() => setNameSortedDesc(false))
      :
      descendingSortedCall("name")
      .then(() => setNameSortedDesc(true))
    :
    (
      unsortAll("name"),
      setNameSorted(true),
      ascendingSortedCall("name")
      .then(() => setNameSortedDesc(false))
    );
  }
  const [isAddressSorted, setAddressSorted] = useState(false);
  const [isAddressSortedDesc, setAddressSortedDesc] = useState(false);
  const onAddressColumnClick = () => {
    isAddressSorted?
      isAddressSortedDesc? 
      ascendingSortedCall("address1_line1")
      .then(() => setAddressSortedDesc(false))
      :
      descendingSortedCall("address1_line1")
      .then(() => setAddressSortedDesc(true))
    :
    (
      unsortAll("address"),
      setAddressSorted(true),
      ascendingSortedCall("address1_line1")
      .then(() => setAddressSortedDesc(false))
    );
  }
  const [isCitySorted, setCitySorted] = useState(false);
  const [isCitySortedDesc, setCitySortedDesc] = useState(false);
  const onCityColumnClick = () => {
    isCitySorted?
      isCitySortedDesc? 
      ascendingSortedCall("address1_city")
      .then(() => setCitySortedDesc(false))
      :
      descendingSortedCall("address1_city")
      .then(() => setCitySortedDesc(true))
    :
    (
      unsortAll("city"),
      setCitySorted(true),
      ascendingSortedCall("address1_city")
      .then(() => setAddressSortedDesc(false))
    );    
  }
  const [isCountrySorted, setCountrySorted] = useState(false);
  const [isCountrySortedDesc, setCountrySortedDesc] = useState(false);
  const onCountryColumnClick = () => {
    isCountrySorted?
      isCountrySortedDesc? 
      ascendingSortedCall("address1_country")
      .then(() => setCountrySortedDesc(false))
      :
      descendingSortedCall("address1_country")
      .then(() => setCountrySortedDesc(true))
    :
    (
      unsortAll("country"),
      setCountrySorted(true),
      ascendingSortedCall("address1_country")
      .then(() => setAddressSortedDesc(false))
    ); 
  }

  const columns: IColumn[] = [
    {
      key: "name",
      name: "Name",
      fieldName: "name",
      minWidth: modalWidth * 0.23,
      maxWidth: modalWidth * 0.27,
      isResizable: true,
      isSorted: isNameSorted,
      isSortedDescending: isNameSortedDesc,
      onColumnClick: onNameColumnClick,
      data: 'string',
      isPadded: true,
    },
    {
      key: "address1_line1",
      name: "Address",
      fieldName: "address1_line1",
      minWidth: modalWidth * 0.23,
      maxWidth: modalWidth * 0.27,
      isSorted: isAddressSorted,
      isSortedDescending: isAddressSortedDesc,
      onColumnClick: onAddressColumnClick,
      isResizable: true,
      data: 'string',
      isPadded: true,
    },
    {
      key: "address1_city",
      name: "City",
      fieldName: "address1_city",
      minWidth: modalWidth * 0.15,
      maxWidth: modalWidth * 0.18,
      isSorted: isCitySorted,
      isSortedDescending: isCitySortedDesc,
      onColumnClick: onCityColumnClick,
      isResizable: true,
      data: 'string',
      isPadded: true,
    },
    {
      key: "address1_country",
      name: "Country",
      fieldName: "address1_country",
      minWidth: modalWidth * 0.15,
      maxWidth: modalWidth * 0.18,
      isSorted: isCountrySorted,
      isSortedDescending: isCountrySortedDesc,
      onColumnClick: onCountryColumnClick,
      isResizable: true,
      data: 'string',
      isPadded: true,
    },
  ];

  //#endregion
  
  //#region Table Data

  const [data, setData] = useState<ComponentFramework.WebApi.Entity[]>([]);
  const [displayedData, setDisplayedData] = useState<
    ComponentFramework.WebApi.Entity[] | null[] | undefined
  >([]);
  const [isDataLoaded, setDataLoaded] = useState(false);
  const [isNextDataLoaded, setNextDataLoaded] = useState(true);
  const selectedTable = "account";
  const queryString = "?$select=name,address1_line1,address1_city,address1_country";
  const FETCH_LENGTH = 60;

  const [nextLink, setNextLink] = useState("");
  const [lastCalledLink, setLastCalledLink] = useState("");

  async function getInitialData() {
    await _context
      .retrieveMultipleRecords(
        selectedTable,
        queryString,
        FETCH_LENGTH
      )
      .then(function (
        response: ComponentFramework.WebApi.RetrieveMultipleResponse
      ) {
        setData(response.entities);
        if (response.nextLink) {
          setNextLink("?" + response.nextLink.split("?")[1]);
        }
      });
  }
  async function getNextPageData() {
    if (nextLink && nextLink !== lastCalledLink) {
      setLastCalledLink(nextLink);
      await _context
        .retrieveMultipleRecords(selectedTable, nextLink, FETCH_LENGTH)
        .then(function (
          response: ComponentFramework.WebApi.RetrieveMultipleResponse
        ) {
          setData(data?.concat(response.entities));
          setNextDataLoaded(true);
          if (response.nextLink) {
            setNextLink("?" + response.nextLink.split("?")[1]);
          } else {
            setNextLink("");
          }
        }).then(() => setTimeout(() => setLoadingNextPage(false),2000));
    }
  }
  async function ascendingSortedCall (column:string) {
    await _context
    .retrieveMultipleRecords(selectedTable,queryString+`&$orderby=${column}%20asc`)
    .then(function (
      response: ComponentFramework.WebApi.RetrieveMultipleResponse
    ) {
      setData(response.entities);
      if (response.nextLink) {
        setNextLink("?" + response.nextLink.split("?")[1]);
      }
    });
  }
  async function descendingSortedCall (column:string) {
    await _context
    .retrieveMultipleRecords(selectedTable,queryString+`&$orderby=${column}%20desc`)
    .then(function (
      response: ComponentFramework.WebApi.RetrieveMultipleResponse
    ) {
      setData(response.entities);
      if (response.nextLink) {
        setNextLink("?" + response.nextLink.split("?")[1]);
      }
    });
  }

  const clearAllData = () => {
     setData([]);
     setDisplayedData([]);
     setDataLoaded(false);
     setNextLink("");
  };
  async function reloadAllData() {
    await _context
      .retrieveMultipleRecords(
        selectedTable,
        queryString
      )
      .then(function (
        response: ComponentFramework.WebApi.RetrieveMultipleResponse
      ) {
        setData(response.entities);
      }).then(() => setDataLoaded(true));
  }

  useEffect(() => {
    getInitialData();
  }, []);

  useEffect(() => {
    setDisplayedData(data);

  }, [data]);


  useEffect(() => {
    if (nextLink) {
      var filtered = data.filter(Boolean);
      //@ts-ignore
      filtered.push(null);
      setDisplayedData(filtered);
      setDataLoaded(true);
    }
  }, [nextLink]);

  //#endregion

  //#region Panel
  const [panelData, setPanelData] =
    useState<ComponentFramework.WebApi.Entity>();
  const [panelAccountId, setPanelAccountId] = useState<string>("");
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    useBoolean(false);
  const [error, setError] = useState(false);
  const [isSaved, setSaved] = useState(false);
  const [callInProgress, setCallInProgress] = useState(false);
  const [isNew, setNew] = useState(false);
  const buttonStyles = { root: { marginRight: 8, marginLeft: 8 } };

  const onItemInvoked = (item: ComponentFramework.WebApi.Entity): void => {
    //alert(`Item Invoked:${JSON.stringify(item)}`)
    setPanelData(item);
    setPanelAccountId(item.accountid.toString());
    openPanel();
  };

  useEffect(() => {
    setUpdatedData({
      ...updatedData,
      name: panelData?.name,
      address1_line1: panelData?.address1_line1,
      address1_city: panelData?.address1_city,
      address1_country: panelData?.address1_country,
    });
  }, [panelData]);

  useEffect(() => {
    console.log(panelAccountId);
  }, [panelAccountId]);

  const [updatedData, setUpdatedData] =
    useState<ComponentFramework.WebApi.Entity>({
      name: "",
      address1_line1: "",
      address1_city: "",
      address1_country: "",
    });
  const handleNameChange = (event: any) => {
    setUpdatedData({ ...updatedData, name: event.target.value });
  };
  const handleAddressChange = (event: any) => {
    setUpdatedData({ ...updatedData, address1_line1: event.target.value });
  };
  const handleCityChange = (event: any) => {
    setUpdatedData({ ...updatedData, address1_city: event.target.value });
  };
  const handleCountryChange = (event: any) => {
    setUpdatedData({ ...updatedData, address1_country: event.target.value });
  };

  async function updatePanelRecord(
    guid: string,
    data: ComponentFramework.WebApi.Entity
  ) {
    setCallInProgress(true);
    
    await _context
      .updateRecord("account", guid, data)
      .then(
        function success(result) {
          setSaved(true);
          setCallInProgress(false);
        },
        function (error) {
          setError(true);
          setCallInProgress(false);
          console.log(error);
        }
      )
      .then(() => clearAllData())
      .then(() => setTimeout(() => reloadAllData(), 1000))
  }

  const cancelUpdate = () => {
    dismissPanel();
  };
  const onPanelDismissed = () => {
    setPanelData(undefined);
    setUpdatedData({
      name: "",
      address1_line1: "",
      address1_city: "",
      address1_country: "",
    });
    setPanelAccountId("");
    setSaved(false);
    setError(false);
    setNew(false);
  }
  const onRenderFooterContent = useCallback(
    (
      guid: string,
      data: ComponentFramework.WebApi.Entity,
      saved: boolean,
      error: boolean,
      isNew: boolean,
      callInProgress: boolean
    ) => (
      <Stack>
        <StackItem>
          {saved ? <p>Saved</p> : null}
          {error ? <p>Error Updating Data</p> : null}
        </StackItem>
        {isNew ? (
          <PrimaryButton
            disabled={callInProgress}
            onClick={() => newAccount(data)}
            styles={buttonStyles}
          >
            Create
          </PrimaryButton>
        ) : (
          <PrimaryButton
            onClick={() => updatePanelRecord(guid, data)}
            styles={buttonStyles}
            disabled={callInProgress}
          >
            Save
          </PrimaryButton>
        )}
        <PrimaryButton onClick={cancelUpdate} styles={buttonStyles}>
          {saved ? <p>Close</p> : <p>Cancel</p>}
        </PrimaryButton>
      </Stack>
    ),
    [dismissPanel]
  );
  //#endregion

  //#region command bar
  async function newAccount(data: ComponentFramework.WebApi.Entity) {
    setCallInProgress(true);
    _context
      .createRecord("account", data)
      .then((resp) => {
        setPanelAccountId(resp.id);
        setNew(false);
        setCallInProgress(false);
      })
      .then(() => clearAllData())
      .then(() => setTimeout(() => reloadAllData(), 2000));
  }
  const openNewPanel = () => {
    setNew(true);
    openPanel();
  };
  const [isSelected, setisSelected] = useState(false);

  async function deleteSelectedRecords() {
    selectionArray
      ? await selectionArray.forEach((element) => {
          console.log(element);
          _context.deleteRecord("account", element);
        })
      : null;
  }

  function onDeleteButtonClick() {
    setisSelected(false);
    deleteSelectedRecords()
      .then(() => clearAllData())
      .then(() => setTimeout(() => reloadAllData(), 1000));
  }
  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: "newItem",
      text: "New",
      iconProps: { iconName: "Add" },
      onClick: () => openNewPanel(),
    },
    {
      key: "deleteItem",
      text: "Delete",
      disabled: !isSelected,
      iconProps: { iconName: "Delete" },
      onClick: onDeleteButtonClick,
    },
  ];
  const [customSearch, setCustomSearch] = useState("");
  const handleCustomSearchChange = (event: any) => {
    setCustomSearch(event.target.value);
  };
  async function getCustomResults(query:string){
    query.length > 1?
      await _context.retrieveMultipleRecords(selectedTable,
        queryString + 
        `&$filter=contains(name,'${query}') or contains(address1_line1,'${query}') or contains(address1_city,'${query}') or contains(address1_country,'${query}')`)
      .then(function (
        response: ComponentFramework.WebApi.RetrieveMultipleResponse
      ) {
        setData(response.entities);
        if (response.nextLink) {
          setNextLink("?" + response.nextLink.split("?")[1]);
        }
      })
      :
      reloadAllData();

  }
  useEffect(() => {
    const timeoutId = setTimeout(() => getCustomResults(customSearch),500);
    return () => clearTimeout(timeoutId);
  },[customSearch])

  //#endregion

  //#region Shimmer

  const [loadingNextPage, setLoadingNextPage] = useState(false);
  const customPlaceHolderCode = () => {
    console.log("called custom code from render");
    setLoadingNextPage(true);
  }
  useEffect(() => {
    if(loadingNextPage)
    {
      getNextPageData();
      console.log("loading next page")
    }
  },[loadingNextPage])

  const onRenderCustomPlaceholder = React.useCallback(
    (
      rowProps: IDetailsRowProps,
      index?: number,
      defaultRender?: (props: IDetailsRowProps) => React.ReactNode
    ): React.ReactNode => {
        //Add code here get data
        displayedData && displayedData.length > 1
          ? setTimeout(() => customPlaceHolderCode(),2000)
          : null;
      return defaultRender ? defaultRender(rowProps) : null;
    },
    [displayedData]
  );
  //#endregion

  //#region Select
  var _selection: Selection;
  _selection = new Selection({
    onSelectionChanged: () => {
      setSelectionArray(_getSelectionIds());
      _selection?.getSelectedCount() > 0 ? 
      setisSelected(true) : setisSelected(false);
    },
  });

  const [selectionArray, setSelectionArray] = useState<string[]>();
  function _getSelectionIds(): string[] {
    const selectionCount = _selection?.getSelectedCount();
    var idArray: string[] = [];
    for (let i = 0; i < selectionCount; i++) {
      idArray.push(
        (_selection?.getSelection()[i] as ComponentFramework.WebApi.Entity)
          .accountid
      );
    }
    return idArray;
  }
  
  return (
    <>
      <Panel
        isOpen={isOpen}
        onDismiss={dismissPanel}
        closeButtonAriaLabel="Close"
        onDismissed={onPanelDismissed}
        onRenderFooterContent={() =>
          onRenderFooterContent(
            panelAccountId,
            updatedData,
            isSaved,
            error,
            isNew,
            callInProgress
          )
        }
        isFooterAtBottom={true}
      >
        <Stack>
          {!isNew ? (
            <TextField readOnly defaultValue={panelAccountId}></TextField>
          ) : null}
          <Stack.Item grow>
            <TextField
              label="Name"
              onChange={handleNameChange}
              defaultValue={updatedData.name}
            />
          </Stack.Item>
          <Stack.Item grow>
            <TextField
              label="Address"
              onChange={handleAddressChange}
              defaultValue={updatedData.address1_line1}
            />
          </Stack.Item>
          <Stack.Item grow>
            <TextField
              label="City"
              onChange={handleCityChange}
              defaultValue={updatedData.address1_city}
            />
          </Stack.Item>
          <Stack.Item grow>
            <TextField
              label="Country"
              onChange={handleCountryChange}
              defaultValue={updatedData.address1_country}
            />
          </Stack.Item>
          {callInProgress ? (
            <Stack.Item grow>
              <Spinner size={SpinnerSize.large} label="Loading..." />
            </Stack.Item>
          ) : null}
        </Stack>
      </Panel>
      <div ref={modalRef}>
        <ScrollablePane 
        scrollbarVisibility={ScrollbarVisibility.auto}
        //initialScrollPosition={1}
        >
          <Sticky stickyPosition={StickyPositionType.Header}>
            <div style={{display:"flex"}}>
              <div style={{width:"50%"}}>
                <CommandBar items={commandBarItems} />
              </div>
              <div style={{width:"50%",margin:"6px"}}>
                <TextField 
                value={customSearch} 
                onChange={handleCustomSearchChange}
                placeholder="Search..."/>
              </div>
            </div>
          </Sticky>
          <MarqueeSelection selection={_selection}>
            <ShimmeredDetailsList 
              setKey="displayedData"
              items={displayedData || []}
              columns={columns}
              enableShimmer={!isDataLoaded}
              flexMargin={20}
              selectionMode={SelectionMode.multiple}
              selection={_selection}
              selectionPreservedOnEmptyClick={true}
              enterModalSelectionOnTouch={true}
              onRenderDetailsHeader={onRenderDetailsHeader}
              onRenderCustomPlaceholder={onRenderCustomPlaceholder}
              onItemInvoked={onItemInvoked}
              ariaLabelForShimmer="Content is being fetched"
              ariaLabelForGrid="Item details"
            />
            {!isNextDataLoaded ? <Spinner size={SpinnerSize.large} /> : null}
          </MarqueeSelection>
        </ScrollablePane>
      </div>
    </>
  );
};
export default ShimmerListComponent;
