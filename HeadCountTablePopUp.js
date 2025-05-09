import { CModal, CModalBody, CModalHeader, CModalTitle } from "@coreui/react";
import React, { useEffect, useState } from "react";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import "primeicons/primeicons.css";
import "primereact/resources/themes/lara-light-indigo/theme.css";
import "primereact/resources/primereact.css";
import "primeflex/primeflex.css";
import "primeicons/primeicons.css";
import { RiFileExcel2Line } from "react-icons/ri";
import { InputText } from "primereact/inputtext";
import { saveAs } from "file-saver";
import ExcelJS from "exceljs";
import { FilterMatchMode } from "primereact/api";
import axios from "axios";
import { environment } from "../../environments/environment";
import moment from "moment";
import Loader from "../Loader/Loader";
import "../RevenueMetrices/HeadCountTableComponent.scss";
function HeadCountTablePopUp(props) {
  const {
    numberLink,
    setNumberLink,
    kpiName,
    totalCount,
    reportRunId,
    selectedMonth,
    searchVal,
    setLoader,
    loader,
    handleAbort,
  } = props;
  const [globalFilterValue, setGlobalFilterValue] = useState("");
  const [filters, setFilters] = useState({
    global: {
      value: null,
      matchMode: FilterMatchMode.CONTAINS,
    },
  });
  const [data, setData] = useState("");
  const baseUrl = environment.baseUrl;
  const onGlobalFilterChange = (e) => {
    const value = e.target.value;
    let _filters = { ...filters };
    _filters["global"].value = value;

    setFilters(_filters);
    setGlobalFilterValue(value);
  };

  const getResourceDetails = () => {
    setLoader(true);
    let name;
    if (kpiName == "Head Count") {
      name = "HC";
    } else if (kpiName == "New Hires") {
      name = "NH";
    } else {
      name = kpiName;
    }
    let searchvalue;
    if (searchVal == null) {
      searchvalue = "-1";
    } else {
      searchvalue = searchVal;
    }
    let month = moment(selectedMonth).format("yyyy-MM-DD");
    axios({
      method: "get",
      url:
        baseUrl +
        `/revenuemetricsms/headCountAndTrend/getResourceDetails?reportRunId=${reportRunId}&selVal=${name}&selMonth=${month}&searchVal=${searchvalue}`,
    }).then((res) => {
      setData(res.data);
      setLoader(false);
    });
  };

  useEffect(() => {
    getResourceDetails();
  }, []);
  const exportExcel = () => {
    const dataOrder = [
      "emp_no",
      "name",
      "cadre",
      "start_date",
      "end_date",
      "is_active",
      "res_cost",
      "revenue",
      "GM_Per",
    ];
    const renamedKeys = {
      emp_no: "Emp ID",
      name: "Emp Name",
      cadre: "Cadre",
      start_date: "Date of Joining",
      end_date: "Last Working Date",
      is_active: "Is Active",
      res_cost: "Cost ($)",
      revenue: "Revenue ($)",
      GM_Per: "Gross Margin %",
    };
    const dataRows = data.map((item) => {
      const row = [];
      dataOrder.forEach((key) => {
        if (renamedKeys[key]) {
          if (key === "is_active") {
            row.push(item[key] == 0 ? "No" : "Yes");
          } else if (key === "start_date" || key === "end_date") {
            row.push(
              item[key] === "" || item[key] === null
                ? ""
                : moment(item[key]).format("DD-MMM-YYYY")
            );
          } else {
            row.push(item[key] || "");
          }
        } else {
          row.push(item[key] || "");
        }
      });
      return row;
    });
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("HeadCount and Margin Trend");

    const headerRow = dataOrder.map((key) => renamedKeys[key] || key);
    worksheet.addRow(headerRow).font = { bold: true };
    dataRows.forEach((row) => {
      worksheet.addRow(row);
    });
    workbook.xlsx.writeBuffer().then((buffer) => {
      saveAs(new Blob([buffer]), "HeadCount and Margin Trend.xlsx");
    });
  };

  const header = () => {
    return (
      <div className="primeTableSearch SUBK-text-n-primesearch-wrapper">
        <div className="Total-NH-SUBK-text">
          Total {(kpiName === "FTE" || kpiName === "SUBK") && "Exits"} {kpiName}{" "}
          : {totalCount}
        </div>
        <div className="primeTableSearch-filter primeTableSearch-filter-Headcount-and-Margins-Trend d-flex gap-2 align-items-center ">
          <span>
            <span className="pi pi-search"></span>
            <InputText
              className="globalFilter"
              placeholder="Keyword Search"
              value={globalFilterValue}
              onChange={onGlobalFilterChange}
            />
          </span>
          <RiFileExcel2Line
            size="2em"
            title="Export to Excel"
            style={{ color: "green" }}
            cursor="pointer"
            onClick={exportExcel}
          />
        </div>
      </div>
    );
  };

  const handleIsActive = (data) => {
    const IsActive = data.is_active === "1" ? "Yes" : "No";
    return (
      <div data-toggle="tooltip" title={IsActive}>
        {IsActive}
      </div>
    );
  };
  const handleDateofJoining = (data) => {
    const date = moment(data.start_date).format("DD-MMM-YYYY");
    return (
      <div data-toggle="tooltip" title={date}>
        {date}
      </div>
    );
  };

  const handleLastWorkingDate = (data) => {
    let date = data.end_date;
    if (date == "" || date == null) {
      date = "";
    } else {
      date = moment(data.end_date).format("DD-MMM-YYYY");
    }

    return (
      <div data-toggle="tooltip" title={date}>
        {date}
      </div>
    );
  };

  const handleCost = (data) => {
    let Value = Number(data.res_cost);
    const formattedValue = Value.toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });

    return (
      <div data-toggle="tooltip" title={formattedValue}>
        {formattedValue}
      </div>
    );
  };
  const handleRecRevenue = (data) => {
    let Value = Number(data.revenue);
    const formattedValue = Value.toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });

    return (
      <div data-toggle="tooltip" title={formattedValue}>
        {formattedValue}
      </div>
    );
  };
  const handleGMPer = (data) => {
    let Value = Number(data.GM_Per);
    const formattedValue = Value.toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });

    return (
      <div data-toggle="tooltip" title={formattedValue}>
        {formattedValue}
      </div>
    );
  };
  const handleCadre = (data) => {
    return (
      <div data-toggle="tooltip" title={data.cadre}>
        {data.cadre === "" || data.cadre === null ? "-" : data.cadre}
      </div>
    );
  };
  const handleEmpName = (data) => {
    return (
      <div data-toggle="tooltip" title={data.name}>
        {data.name}
      </div>
    );
  };
  const handleEmpId = (data) => {
    return (
      <div data-toggle="tooltip" title={data.emp_no}>
        {data.emp_no}
      </div>
    );
  };

  const columns = [
    {
      field: "S.No",
      header: "S.No",
      body: (data, options) => (
        <span data-toggle="tooltip" title={options.rowIndex + 1}>
          {options.rowIndex + 1}
        </span>
      ),
      bodyStyle: { textAlign: "center" },
    },
    { field: "emp_no", header: "Emp. ID", body: handleEmpId },
    { field: "name", header: "Resource", body: handleEmpName },
    {
      field: "cadre",
      header: "Cadre",
      body: handleCadre,
      bodyStyle: { textAlign: "center" },
    },
    {
      field: "dept",
      header: "Business Unit",
      body: (data) => (
        <span data-toggle="tooltip" title={data.dept}>
          {data.dept}
        </span>
      ),
      bodyStyle: { textAlign: "left" },
    },
    {
      field: "start_date",
      header: "DOJ",
      body: handleDateofJoining,
      bodyStyle: { textAlign: "center" },
    },
    {
      field: "end_date",
      header: "LWD",
      body: handleLastWorkingDate,
      bodyStyle: { textAlign: "center" },
    },
    {
      field: "is_active",
      header: "Is Active",
      body: handleIsActive,
      bodyStyle: { textAlign: "center" },
    },
    {
      field: "res_cost",
      header: "Cost ($)",
      body: handleCost,
      bodyStyle: { textAlign: "right" },
    },
    {
      field: "revenue",
      header: "Rec. Rev ($)",
      body: handleRecRevenue,
      bodyStyle: { textAlign: "right" },
    },
    {
      field: "GM_Per",
      header: "GM (%)",
      body: handleGMPer,
      bodyStyle: { textAlign: "right" },
    },
  ];
  return (
    <div>
      <CModal
        size="xl"
        alignment="center"
        backdrop="static"
        className="ui-dialog"
        visible={numberLink}
        onClose={() => setNumberLink(false)}
      >
        <CModalHeader className="hgt22">
          <CModalTitle className="ft16">{selectedMonth}</CModalTitle>
        </CModalHeader>
        <CModalBody>
          <DataTable
            value={data}
            showGridlines
            dataKey="id"
            className="primeReactDataTable darkHeader "
            stripedRows
            paginatorTemplate="RowsPerPageDropdown FirstPageLink PrevPageLink CurrentPageReport NextPageLink LastPageLink"
            currentPageReportTemplate="{first} to {last} of {totalRecords}"
            paginator
            pagination="true"
            rows={15}
            paginationPerPage={5}
            rowsPerPageOptions={[10, 25, 50]}
            sortMode="multiple"
            emptyMessage="No Data Found"
            filters={filters}
            header={header}
          >
            {columns.map((col, index) => (
              <Column
                key={index}
                field={col.field}
                header={
                  <span data-toggle="tooltip" title={col.header}>
                    {col.header}
                  </span>
                }
                sortable
                body={col.body}
                bodyStyle={col.bodyStyle}
              />
            ))}
          </DataTable>
        </CModalBody>
      </CModal>
      {loader ? <Loader handleAbort={handleAbort} /> : ""}
    </div>
  );
}
export default HeadCountTablePopUp;
