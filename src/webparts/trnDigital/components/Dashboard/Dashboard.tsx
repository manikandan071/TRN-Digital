import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web, ICamlQuery } from "@pnp/sp/presets/all";
import styles from "../TrnDigital.module.scss";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  IIconProps,
  SelectionMode,
  IconButton,
  DefaultButton,
  Modal,
  Dropdown,
  IDropdownStyles,
  TextField,
} from "@fluentui/react";
import Pagination from "@material-ui/lab/Pagination";
import CustomLoader from "../Loder/CustomLoder";
import { sp } from "sp-pnp-js";
// import { ICamlQuery } from "@pnp/sp/presets/all";

export default function Dashboard(prop: any): JSX.Element {
  let spweb = Web("https://technorucs365.sharepoint.com/sites/DemoTrnDigital");
  // let sharepointWeb = Web(prop.URL);
  let currpage = 1;
  let totalPageItems = 30;
  let dropdownOptions = [
    { key: "File", text: "File" },
    { key: "DrawingNumber", text: "DrawingNumber" },
    { key: "DocumentTitle", text: "DocumentTitle" },
    { key: "DocumentDate", text: "DocumentDate" },
    { key: "DocumentDescription", text: "DocumentDescription" },
    { key: "RevisionNumber", text: "RevisionNumber" },
    { key: "VendorName", text: "VendorName" },
  ];
  let allData = [];

  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          alignItems: "start",
          ".ms-DetailsRow-cell": {
            display: "flex",
            alignItems: "center",
            height: 50,
            minHeight: 50,
            padding: "5px 10px",
            margin: "auto",
            color: "#000",
          },
          // ".ms-DetailsHeader": {
          //   background: "#000",
          // },
          ".ms-DetailsHeader-cellName": {
            color: "#fff",
          },
          ".ms-DetailsHeader-cellTitle": {
            padding: "0px 8px 0px 10px",
            background: "#446f3c",
          },
        },
        ".root-154": {
          color: "#f0d8d8",
          backgroundColor: "#3635399e",
        },
        ".root-140": {
          borderBottom: "1px solid #b8bbde",
        },
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };

  const modalStyle = {
    root: {
      ".ms-Dialog-main": {
        width: "55%",
        height: "auto",
      },
    },
  };

  const dropdownStyles: Partial<IDropdownStyles> = {
    root: { width: "20%", marginRight: "22px" },
    dropdown: { width: "100%" },
  };

  const View: IIconProps = { iconName: "View" };
  const Edit: IIconProps = { iconName: "Edit" };
  const Close: IIconProps = { iconName: "ChromeClose" };
  const [loader, setLoader] = useState(false);
  const [masterData, setMasterData] = useState([]);
  const [displayData, setDisplayData] = useState([]);
  const [duplicateData, setDuplicateData] = useState([]);
  const [selectList, setSelectList] = useState([]);
  const [selectOption, setSelectOption] = useState("File");
  const [searchValue, setSearchValue] = useState("");
  const [currentPage, setCurrentPage] = useState(currpage);
  const [IsEditPopup, setIsEditPopup] = useState(false);

  let columns = [
    {
      key: "columns1",
      name: "File",
      fieldName: "File Name",
      minWidth: 150,
      maxWidth: 190,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.File}</div>
        </>
      ),
    },
    {
      key: "columns2",
      name: "Drawing Number",
      fieldName: "DrawingNumber",
      minWidth: 150,
      maxWidth: 190,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.DrawingNumber ? item.DrawingNumber : "-"}</div>
        </>
      ),
    },
    {
      key: "columns3",
      name: "Document Title",
      fieldName: "DocumentTitle",
      minWidth: 150,
      maxWidth: 190,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.DocumentTitle ? item.DocumentTitle : "-"}</div>
        </>
      ),
    },
    {
      key: "columns4",
      name: "Document Description",
      fieldName: "DocumentDescription",
      minWidth: 150,
      maxWidth: 190,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.DocumentDescription ? item.DocumentDescription : "-"}</div>
        </>
      ),
    },
    {
      key: "columns5",
      name: "Document Date",
      fieldName: "DocumentDate",
      minWidth: 80,
      maxWidth: 100,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.DocumentDate ? item.DocumentDate : "-"}</div>
        </>
      ),
    },
    {
      key: "columns6",
      name: "Revision Number",
      fieldName: "RevisionNumber",
      minWidth: 150,
      maxWidth: 190,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.RevisionNumber ? item.RevisionNumber : "-"}</div>
        </>
      ),
    },
    {
      key: "columns7",
      name: "Vendor Name",
      fieldName: "VendorName",
      minWidth: 150,
      maxWidth: 190,
      isMultiline: true,
      onRender: (item) => (
        <>
          <div>{item.VendorName ? item.VendorName : "-"}</div>
        </>
      ),
    },
    {
      key: "columns8",
      name: "Actions",
      fieldName: "FileUrl",
      minWidth: 80,
      maxWidth: 100,
      onRender: (item) => (
        <>
          <div>
            <IconButton
              className={styles.viewBtn}
              iconProps={View}
              style={{ cursor: "pointer", marginRight: "10px", fontSize: 72 }}
              title="View File"
              ariaLabel="View File"
              onClick={
                (ev) => openDocumentFileFun(item, item.FileUrl)
                //   setIsEditPopup(true)
              }
            />
            <IconButton
              className={styles.editBtn}
              iconProps={Edit}
              style={{ cursor: "pointer", fontSize: 72 }}
              title="Edit"
              ariaLabel="Edit"
              onClick={(ev) => (editFuntion(item), setIsEditPopup(true))}
            />
          </div>
        </>
      ),
    },
  ];
  const getData = () => {
    spweb.lists
      .getByTitle(`Documents`)
      .items.top(5000)
      .select(
        "FileLeafRef",
        "Title",
        "DrawingNumber",
        "VendorName",
        "DocumentDate",
        "DocumentDescription",
        "DocumentTitle",
        "RevisionNumber",
        "ID"
      )
      .orderBy("ID", true)
      //   .expand("Name")
      .get()
      .then((res) => {
        let documentListsArr: any[] = [];
        if (res.length > 0) {
          res.forEach((data) => {
            documentListsArr.push({
              Id: data.ID,
              DrawingNumber: data.DrawingNumber,
              VendorName: data.VendorName,
              DocumentDate: moment(data.DocumentDate).format("YYYY-MM-DD"),
              DocumentDescription: data.DocumentDescription,
              DocumentTitle: data.DocumentTitle,
              RevisionNumber: data.RevisionNumber,
              File: data.FileLeafRef,
            });
          });
          setMasterData([...documentListsArr]);
          setDisplayData([...documentListsArr]);
          setDuplicateData([...documentListsArr]);
          paginateFunction(currentPage, [...documentListsArr]);
        }
      });
  };
  const editFuntion = (item) => {
    setSelectList([item]);
    setIsEditPopup(true);
  };
  const openDocumentFileFun = (id, file) => {
    let Url = "https://technorucs365.sharepoint.com" + file;
    window.open(Url, "_blank");

    // sp.web
    //   .getFolderByServerRelativeUrl(Url)
    //   .files.get()
    //   .then((res) => {
    //     console.log(res, id);
    //   });
  };

  const dropdownSelectFun = (value) => {
    setSelectOption(value);
    setSearchValue("");
    setDuplicateData([...masterData]);
    paginateFunction(1, [...masterData]);
  };

  const onChangeSearchFun = (event, value) => {
    setSearchValue(value);
    let SearchFilter = masterData.filter((data) => {
      if (data[selectOption] && value != "") {
        return data[selectOption].toLowerCase().match(value.toLowerCase());
      } else if (value == "") {
        return data;
      }
      // data[selectOption].toLowerCase().includes(value.toLowerCase())
    });
    if (SearchFilter.length > 0) {
      setDuplicateData([...SearchFilter]);
      paginateFunction(1, [...SearchFilter]);
    } else {
      setDuplicateData([]);
      paginateFunction(1, []);
    }
    // setDisplayData([...SearchFilter]);
    // setDuplicateData([...SearchFilter]);
  };

  const onChangeFun = (value, type) => {
    let editAyy = selectList[0];
    let onchangeObject = {
      Id: editAyy.Id,
      DrawingNumber: editAyy.DrawingNumber,
      VendorName: editAyy.VendorName,
      DocumentDate: editAyy.DocumentDate,
      DocumentDescription: editAyy.DocumentDescription,
      DocumentTitle: editAyy.DocumentTitle,
      RevisionNumber: editAyy.RevisionNumber,
      File: editAyy.File,
    };
    onchangeObject = {
      Id: editAyy.Id,
      DrawingNumber: type === "DiaNo" ? value : editAyy.DrawingNumber,
      VendorName: type === "VendorName" ? value : editAyy.VendorName,
      DocumentDate: type === "date" ? value : editAyy.DocumentDate,
      DocumentDescription:
        type === "DocDescription" ? value : editAyy.DocumentDescription,
      DocumentTitle: type === "DocTitle" ? value : editAyy.DocumentTitle,
      RevisionNumber: type === "Revision" ? value : editAyy.RevisionNumber,
      File: editAyy.File,
    };
    setSelectList([onchangeObject]);
  };

  const submitFunction = async () => {
    try {
      await spweb.lists
        .getByTitle(`${prop.libraryName}`)
        .items.getById(selectList[0].Id)
        .update({
          DrawingNumber: selectList[0].DrawingNumber,
          DocumentDate: selectList[0].DocumentDate + "T08:00:00Z",
          // DocumentDate: "2023-08-11T08:00:00Z",
          DocumentDescription: selectList[0].DocumentDescription,
          DocumentTitle: selectList[0].DocumentTitle,
          RevisionNumber: selectList[0].RevisionNumber,
          VendorName: selectList[0].VendorName,
        })
        .then((res) => {
          masterData.forEach((data) => {
            if (data.Id === selectList[0].Id) {
              data.DrawingNumber = selectList[0].DrawingNumber;
              data.DocumentDate = selectList[0].DocumentDate;
              data.DocumentDescription = selectList[0].DocumentDescription;
              data.DocumentTitle = selectList[0].DocumentTitle;
              data.RevisionNumber = selectList[0].RevisionNumber;
              data.VendorName = selectList[0].VendorName;
            }
          });
          displayData.forEach((data) => {
            if (data.Id === selectList[0].Id) {
              data.DrawingNumber = selectList[0].DrawingNumber;
              data.DocumentDate = selectList[0].DocumentDate;
              data.DocumentDescription = selectList[0].DocumentDescription;
              data.DocumentTitle = selectList[0].DocumentTitle;
              data.RevisionNumber = selectList[0].RevisionNumber;
              data.VendorName = selectList[0].VendorName;
            }
          });
          setIsEditPopup(false);
        });
    } catch (err) {
      console.log(err);
    }
  };

  function getPagedData(data, query) {
    let tempArr = [];
    spweb.lists
      .getByTitle(`${prop.libraryName}`)
      .renderListDataAsStream({ ViewXml: query, Paging: data.substring(1) })
      .then((data) => {
        allData.push(...data.Row);
        if (data.NextHref) {
          getPagedData(data.NextHref, query);
        } else {
          allData.forEach((data) => {
            let fileName = data.FileRef.split("/").pop();
            tempArr.push({
              Id: data.ID,
              DrawingNumber: data.DrawingNumber ? data.DrawingNumber : "",
              VendorName: data.VendorName ? data.VendorName : "",
              DocumentDate: data.DocumentDate
                ? moment(data.DocumentDate).format("YYYY-MM-DD")
                : "",
              DocumentDescription: data.DocumentDescription
                ? data.DocumentDescription
                : "",
              DocumentTitle: data.DocumentTitle ? data.DocumentTitle : "",
              RevisionNumber: data.RevisionNumber ? data.RevisionNumber : "",
              FileUrl: data.FileRef,
              File: fileName,
            });
          });
          setMasterData([...tempArr]);
          setDisplayData([...tempArr]);
          setDuplicateData([...tempArr]);
          paginateFunction(currentPage, [...tempArr]);
          setLoader(false);
        }
      })
      .catch((err) => {
        console.log(err), setLoader(false);
      });
  }

  function getAllData(query) {
    spweb.lists
      .getByTitle(`${prop.libraryName}`)
      .renderListDataAsStream({ ViewXml: query })
      .then((data) => {
        allData.push(...data.Row);
        if (data.NextHref) {
          getPagedData(data.NextHref, query);
        }
      })
      .catch((err) => {
        console.log(err), setLoader(false);
      });
  }

  const fetchItems = async () => {
    setLoader(true);
    let camlQuery = `
          <View Scope='RecursiveAll'>
            <Query>
              <OrderBy>
                <FieldRef Name='ID' Ascending='TRUE'/>
              </OrderBy>
            </Query>
            <ViewFields>
            <FieldRef Name='ID' />
              <FieldRef Name='DrawingNumber' />
              <FieldRef Name='DocumentDate' />
              <FieldRef Name='DocumentDescription' />
              <FieldRef Name='DocumentTitle' />
              <FieldRef Name='RevisionNumber' />
              <FieldRef Name='VendorName' />
            </ViewFields>
            <RowLimit Paged='TRUE'>200</RowLimit>
          </View>`;

    getAllData(camlQuery);
  };

  useEffect(() => {
    // getData();
    fetchItems();
  }, []);

  const paginateFunction = (pagenumber, data: any[]) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currpage = pagenumber;
      setCurrentPage(pagenumber);
      setDisplayData(paginatedItems);
    } else {
      setDisplayData([]);
      // setCurrentPage(1);
    }
  };

  return loader ? (
    <CustomLoader />
  ) : (
    <div style={{ padding: "0px 20px" }}>
      <Modal isOpen={IsEditPopup} styles={modalStyle}>
        <div>
          <div style={{ textAlign: "end", padding: "5px 5px 0px 0px" }}>
            <IconButton
              iconProps={Close}
              style={{
                fontSize: 60,
                cursor: "pointer",
              }}
              title="Close"
              ariaLabel="Close"
              onClick={() => {
                setIsEditPopup(false);
              }}
            />
          </div>
          {selectList.length > 0 ? (
            <div className={styles.modelMain}>
              <div className={styles.modelTitle}>
                <h3
                  style={{
                    margin: "0px",
                    padding: "8px 0px",
                    color: "#fff",
                    fontSize: "15px",
                    fontWeight: "600",
                  }}
                >
                  {selectList[0].File}
                </h3>
              </div>
              <div className={styles.modelSec}>
                <div className={styles.modelBox}>
                  <label style={{ width: "45%", fontWeight: "700" }}>
                    Drawing No :
                  </label>
                  <input
                    style={{ width: "55%", padding: "4px!important" }}
                    type="text"
                    value={selectList[0].DrawingNumber}
                    onChange={(e) => onChangeFun(e.target.value, "DiaNo")}
                  />
                </div>
              </div>
              <div className={styles.modelSec}>
                <div className={styles.modelBox}>
                  <label style={{ width: "45%", fontWeight: "700" }}>
                    Document Date :
                  </label>
                  <input
                    style={{ width: "55%" }}
                    type="date"
                    value={selectList[0].DocumentDate}
                    onChange={(e) => onChangeFun(e.target.value, "date")}
                  />
                </div>
              </div>
              <div className={styles.modelSec}>
                <div className={styles.modelBox}>
                  <label style={{ width: "45%", fontWeight: "700" }}>
                    Document Title :
                  </label>
                  <textarea
                    style={{ width: "55%", padding: "4px" }}
                    // type="text"
                    value={selectList[0].DocumentTitle}
                    onChange={(e) => onChangeFun(e.target.value, "DocTitle")}
                  />
                </div>
              </div>
              <div className={styles.modelSec}>
                <div className={styles.modelBox}>
                  <label style={{ width: "45%", fontWeight: "700" }}>
                    Document Description :
                  </label>
                  <textarea
                    style={{ width: "55%" }}
                    // type="text"
                    value={selectList[0].DocumentDescription}
                    onChange={(e) =>
                      onChangeFun(e.target.value, "DocDescription")
                    }
                  />
                </div>
              </div>
              <div className={styles.modelSec}>
                <div className={styles.modelBox}>
                  <label style={{ width: "45%", fontWeight: "700" }}>
                    Revision Number :
                  </label>
                  <input
                    style={{ width: "55%" }}
                    type="text"
                    value={selectList[0].RevisionNumber}
                    onChange={(e) => onChangeFun(e.target.value, "Revision")}
                  />
                </div>
              </div>
              <div className={styles.modelSec}>
                <div className={styles.modelBox}>
                  <label style={{ width: "45%", fontWeight: "700" }}>
                    Vendor Name :
                  </label>
                  <input
                    style={{ width: "55%" }}
                    type="text"
                    value={selectList[0].VendorName}
                    onChange={(e) => onChangeFun(e.target.value, "VendorName")}
                  />
                </div>
              </div>
              <div className={styles.modelButton}>
                <DefaultButton
                  primary
                  text={"Submit"}
                  style={{
                    cursor: "pointer",
                    backgroundColor: "#67c25f",
                    border: "1px solid #67c25f",
                    marginRight: "20px",
                  }}
                  onClick={() => submitFunction()}
                />
                <DefaultButton
                  primary
                  text={"Cancel"}
                  style={{
                    cursor: "pointer",
                    backgroundColor: "#be3535ed",
                    border: "1px solid #be3535ed",
                  }}
                  onClick={() => {
                    setIsEditPopup(false);
                  }}
                />
              </div>
            </div>
          ) : (
            <div></div>
          )}
        </div>
      </Modal>

      <div>
        <div style={{ display: "flex", paddingBottom: "10px" }}>
          <Dropdown
            label="Select Property"
            selectedKey={selectOption}
            onChange={(e, option) => {
              dropdownSelectFun(option["text"]);
              // filterHandleFunction("status", option["text"]);
            }}
            placeholder="Select an option"
            options={dropdownOptions}
            styles={dropdownStyles}
          />
          <TextField
            // className={styles.textfeild}
            value={searchValue}
            styles={dropdownStyles}
            label="Search"
            onChange={onChangeSearchFun}
            placeholder="Search"
          />
        </div>
        <DetailsList
          items={displayData}
          columns={columns}
          setKey="set"
          styles={gridStyles}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          //   onRenderRow={onRenderRow}
        />
      </div>
      {displayData.length == 0 ? (
        <div className={styles.noRecordsec}>
          <h4>No records found !!!</h4>
        </div>
      ) : (
        <div className={styles.pagination}>
          <Pagination
            page={currentPage}
            onChange={(e, page) => {
              paginateFunction(page, duplicateData);
            }}
            count={
              duplicateData.length > 0
                ? Math.ceil(duplicateData.length / totalPageItems)
                : 1
            }
            color="primary"
            showFirstButton={currentPage == 1 ? false : true}
            showLastButton={
              currentPage == Math.ceil(duplicateData.length / totalPageItems)
                ? false
                : true
            }
          />
        </div>
      )}
    </div>
  );
}
