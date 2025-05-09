import React, { useState, useEffect, useRef } from "react";
import { MultiSelect } from "react-multi-select-component";
import axios from "axios";
import { environment } from "../../environments/environment";
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";
import moment from "moment";
import {
  FaChevronCircleDown,
  FaChevronCircleUp,
  FaSearch,
  FaCaretDown,
} from "react-icons/fa";
import { CCollapse } from "@coreui/react";
import { ReactSearchAutocomplete } from "react-search-autocomplete";
import GlobalValidation from "../ValidationComponent/GlobalValidation";
import HeadCountTableComponent from "./HeadCountTableComponent";
import { AiFillWarning } from "react-icons/ai";
import SelectCustDialogBox from "../Customer/SelectCustDialogBox";
import Loader from "../Loader/Loader";
import GlobalHelp from "../PrimeReactTableComponent/GlobalHelp";
import Utils from "../../Utils";
import { logInfo, logWarn, logError, logDebug } from "../../logger/logger";
import ScreenBreadcrumbs from "../Common/ScreenBreadcrumbs";

function HeadCount() {
  const abortController = useRef(null);
  const [loader, setLoader] = useState(false);
  const [departments, setDepartments] = useState([]);
  const [selectedDepartments, setSelectedDepartments] = useState([]);
  const [country, setCountry] = useState([]);
  const [selectedCountry, setSelectedCountry] = useState([]);
  const [cadres, setCadres] = useState([]);
  const [selectedCadres, setSelectedCadres] = useState([]);
  const [measures, setMeasures] = useState([]);
  const [selectedMeasures, setSelectedMeasures] = useState([]);
  const [IsSearched, setIsSearched] = useState(false);

  const findLaterMeasureLabel = (measures, selectedMeasures) => {
    const selectedLabels = selectedMeasures.map((measure) => measure.label);
    const laterMeasures = measures.filter((measure) =>
      selectedLabels.includes(measure.label)
    );
    const lastLaterMeasure = laterMeasures[laterMeasures.length - 1];
    return lastLaterMeasure ? lastLaterMeasure.label : "";
  };

  const laterMeasureLabel = findLaterMeasureLabel(measures, selectedMeasures);

  const [csl, setCsl] = useState([]);
  const [selectedCsl, setSelectedCsl] = useState([]);
  const [delivery, setDelivery] = useState([]);
  const [selectedDelivery, setSelectedDelivery] = useState([]);
  const [accountOwner, setAccountOwner] = useState([]);
  const [engComp, setEngComp] = useState([]);
  const [selectedAccountOwner, setSelectedAccountOwner] = useState([]);
  const [selectedEngComp, setSelectedEngComp] = useState([]);
  const [visible, setVisible] = useState(false);
  const [cheveronIcon, setCheveronIcon] = useState(FaChevronCircleUp);
  const [project, setProject] = useState([]);
  const [customer, setCustomer] = useState([]);
  const [tableData, setTableData] = useState([]);
  const [column, setColumn] = useState([]);
  const [selectType, setSelectType] = useState("BusinessUnit");
  const baseUrl = environment.baseUrl;
  const [month, setMonth] = useState(
    new Date().setMonth(new Date().getMonth() - 3)
  );
  const currentDate = new Date();
  const selectedMonth = new Date(month);
  const diffMonths =
    (currentDate.getFullYear() - selectedMonth.getFullYear()) * 12 +
    currentDate.getMonth() -
    selectedMonth.getMonth();

  const loggedUserId = localStorage.getItem("resId");
  const [validationmessage, setValidationMessage] = useState(false);
  const ref = useRef([]);
  const [displayTableState, setDisplayTableState] = useState(false);
  const [custVisible, setCustVisible] = useState(false);
  const [toggleButton, setToggleButton] = useState(false);
  const [reportRunId, setReportRunId] = useState("");
  const [routes, setRoutes] = useState([]);
  let textContent = "Revenue Metrics";
  let currentScreenName = ["Headcount and Margins Trend"];
  sessionStorage.setItem(
    "breadCrumbs",
    JSON.stringify({
      routes: routes,
      currentScreenName: currentScreenName,
      textContent: textContent,
    })
  );

  useEffect(() => {
    getMenus();
  }, []);

  const getMenus = () => {
    axios({
      method: "GET",
      url: baseUrl + `/CommonMS/master/getMenus?loggedUserId=${loggedUserId}`,
    }).then((resp) => {
      logInfo("Fetched Menus Succesfully","getMenus");
      const modifiedUrlPath = "/pmo/headCountTrend";
      getUrlPath(modifiedUrlPath);
      let getData = resp.data.map((menu) => {
        if (menu.subMenus) {
          menu.subMenus = menu.subMenus.filter(
            (subMenu) =>
              subMenu.display_name !== "Rev. Projections" &&
              subMenu.display_name !== "Project Timesheet (Deprecated)" &&
              subMenu.display_name !== "Financial Plan & Review"
          );
        }
        return menu;
      });
      const updatedMenuData = resp.data.map((category) => ({
        ...category,
        subMenus: category.subMenus.map((submenu) => {
          if (submenu.display_name === "Revenue By Industry") {
            return {
              ...submenu,
              display_name: "Recognized Revenue By Industry",
            };
          }

          return submenu;
        }),
      }));
      updatedMenuData.forEach((item) => {
        if (item.display_name === textContent) {
          setRoutes([item]);
          sessionStorage.setItem("displayName", item.display_name);
        }
      });
    })
    .catch((error) => {
      logError("Failed to fetch Menus", "getMenus");
    });
  };
  const getUrlPath = (modifiedUrlPath) => {
    axios({
      method: "get",
      url:
        baseUrl +
        `/CommonMS/security/authorize?url=${modifiedUrlPath}&userId=${loggedUserId}`,
    })
      .then((res) => {logInfo("Authorization successful", "getUrlPath");})
      .catch((error) => {logError("Authorization Not successful", "getUrlPath");});
  };

  const ArrowRenderer = ({ expanded }) => (
    <>
      {expanded ? (
        <FaCaretDown className="chevronIcon" />
      ) : (
        <FaCaretDown className="chevronIcon" />
      )}
    </>
  );

  const Data = moment(month).startOf("month").format("yyyy-MM-DD");

  const initialValue = {
    month: Data,
    duration: 3,
    resTyp: -1,
    countries: "",
    measures: "-1",
    searchType: "BusinessUnit",
    busUnits: "170,211,123,82,168,207,212,18,213,49,149,208,999",
    Customer: -1,
    prjId: -1,
    sortBy: -1,
    cadres: -1,
    cslRes: -1,
    dpRes: -1,
    objId: -1,
    accOwnerRes: -1,
    overheadChecked: false,
    engComp: -1,
  };
  const [searchdata, setSearchdata] = useState(initialValue);
  const handleMonthChange = (date) => {
    setMonth(date);
    const durationSelect = document.getElementById("duration");
    durationSelect.value = "";
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    if (name === "searchType") {
      setSelectType(value);
      setDisplayTableState(false);
      if (value === "BusinessUnit") {
        setSearchdata((prev) => ({
          ...prev,
          cslRes: -1,
          dpRes: -1,
          prjId: -1,
          Customer: -1,
          accOwnerRes: -1,
        }));
      } else if (value === "Customer") {
        setSearchdata((prev) => ({
          ...prev,
          cslRes: -1,
          dpRes: -1,
          prjId: -1,
          busUnits: -1,
          Customer: 0, // Set the busUnits property to the correct value
          accOwnerRes: -1,
        }));
      } else if (value === "CSL") {
        setSearchdata((prev) => ({
          ...prev,
          busUnits: -1,
          dpRes: -1,
          prjId: -1,
          Customer: -1,
          accOwnerRes: -1,
        }));
      } else if (value === "Project") {
        setSearchdata((prev) => ({
          ...prev,
          cslRes: -1,
          dpRes: -1,
          busUnits: -1,
          Customer: -1,
          accOwnerRes: -1,
        }));
      } else if (value === "AccountOwner") {
        setSearchdata((prev) => ({
          ...prev,
          cslRes: -1,
          dpRes: -1,
          busUnits: -1,
          Customer: -1,
          prjId: -1,
        }));
      } else {
        setSearchdata((prev) => ({
          ...prev,
          cslRes: -1,
          busUnits: -1,
          prjId: -1,
          Customer: -1,
          accOwnerRes: -1,
        }));
      }
    }

    if (name == "Customer" && value === "1") {
      setCustVisible(true);
    }
    setSearchdata((prev) => {
      return { ...prev, [name]: value };
    });
  };

  const [selectedItems, setSelectedItems] = useState([{}]);
  const Customer = selectedItems?.map((d) => d?.id).toString();

  useEffect(() => {}, [Customer]);
  const selectedCust = JSON.parse(localStorage.getItem("selectedCust"))
    ?.map((d) => d.id)
    ?.toString();

  const customValueRenderer = (selected, _options) => {
    return selected.length === country.length ||
      selected.length === csl.length ||
      selected.length === delivery.length
      ? "<< ALL >>"
      : selected.length === 0
      ? "<< Please Select >>"
      : selected.map((label) => {
          return selected.length > 1 ? label.label + "," : label.label;
        });
  };
  const handleSearch = (e) => {
    setValidationMessage(false);
    abortController.current = new AbortController();

    let filteredData = ref.current.filter((d) => d != null);
    ref.current = filteredData;
    let valid = GlobalValidation(ref);

    if (valid == true) {
      setLoader(false);
      setValidationMessage(true);
      return;
    }

    setDisplayTableState(false);
    const loaderTime = setTimeout(() => {
      setLoader(true);
    }, 2000);
    axios({
      method: "post",
      url:
        baseUrl +
        `/revenuemetricsms/headCountAndTrend/getHeadCountAndTrendData`,
      signal: abortController.current.signal,
      data: {
        month: Data,
        duration: searchdata.duration,
        resTyp: searchdata.resTyp,
        countries:
          searchdata.countries == "6,5,3,8,0,7,1,2"
            ? "6,5,3,8,7,1,2"
            : searchdata.countries,
        measures: searchdata.measures,
        searchType: searchdata.searchType,
        busUnits:
          searchdata.searchType != "BusinessUnit" ? -1 : searchdata.busUnits,
        customers:
          searchdata.searchType == "Customer"
            ? searchdata.Customer == "1"
              ? selectedCust
              : searchdata.Customer == "-1"
              ? "-1"
              : 0
            : -1,
        prjId: searchdata.prjId,
        sortBy: searchdata.sortBy,
        cadres:
          searchdata.cadres == "E1,E2,E3,L,M,G,T1,UA" ? -1 : searchdata.cadres,
        cslRes: searchdata.cslRes,
        dpRes: searchdata.dpRes,
        objId: searchdata.objId, // change in tables according to the expandable we choose
        accOwnerRes:
          searchdata.searchType == "AccountOwner"
            ? searchdata.accOwnerRes
            : searchdata.accOwnerRes ==
              accountOwner.map((item) => item.value).join(",")
            ? -1
            : searchdata.accOwnerRes,
        overheadChecked: searchdata.overheadChecked,
        engComp: searchdata.searchType == "engComp" ? searchdata.engComp : -1,
      },
    }).then((res) => {
      {logInfo("Authorization_successful", "handleSearch");}
      setLoader(false);
      clearTimeout(loaderTime);
      let respData = res.data;
      respData.data.forEach((item) => {
        item.uniqueId = item.id;
      });
      setTableData(respData.data);
      setToggleButton(true);
      setReportRunId(respData.reportRunId);
      axios({
        method: "get",
        url:
          baseUrl +
          `/revenuemetricsms/headCountAndTrend/getColumnsData?reportRunId=${respData.reportRunId}`,
        signal: abortController.current.signal,
      })
        .then((res1) => {
          {logInfo("Authorizationsuccessful_ReportRunId", "handleSearch");}
          !valid && setVisible(!visible);
          visible
            ? setCheveronIcon(FaChevronCircleUp)
            : setCheveronIcon(FaChevronCircleDown);
          const responseString1 = res1.data.val.replace(/`/g, "");
          const responseString2 = responseString1.replace(/'/g, "");
          const dataArray = responseString2.split(",");
          setColumn(dataArray.slice(1));
          setIsSearched(true);
        })
        .catch((error) => {logError("Authorization Not successful", "handleSearch");});

      setValidationMessage(false);
      setDisplayTableState(true);
    })
    .catch((error) => {
      logError("Authorization Not successful", "handleSearch");
    });
  };

  const handleAbort = () => {
    abortController.current && abortController.current.abort();
    setLoader(false);
  };
  const getDepartments = async () => {
    const resp = await axios({
      url: baseUrl + `/CostMS/cost/getDepartments`, //13
    });

    let departments = resp.data;
    departments = departments.filter((ele) => ele.value >= 0);
    departments.push({ value: 999, label: "Non-Revenue Units" });
    setDepartments(departments);
    setSelectedDepartments(departments.filter((ele) => ele.value >= 0));
    let filteredDeptData = [];
    departments.forEach((data) => {
      filteredDeptData.push(data.value);
    });
    setSearchdata((prevVal) => ({
      ...prevVal,
      ["busUnits"]: filteredDeptData.toString(),
    }));
  };
  const handleCsl = () => {
    const loggedUser = loggedUserId;
    axios({
      method: "get",
      url:
        baseUrl +
        `/SalesMS/MasterController/getCslDropDownData?userId=${loggedUser}`,
    }).then((res) => {
      {logInfo("CslDropDownData Fetched", "handleCsl");}
      let custom = [];
      let data = res.data;
      data.length > 0 &&
        data.forEach((e) => {
          let cslObj = {
            label: e.PersonName,
            value: e.id,
          };
          custom.push(cslObj);
        });
      setCsl(custom);
      setSelectedCsl(custom);
    })
    .catch((error) => {
      logError("Failed to fetch CslDropDownData", "handleCsl");
    });
  };
  const getCustomerdata = () => {
    axios({
      method: "get",
      url: baseUrl + `/revenuemetricsms/RevenueMarginAnalysis/getCustomers`,
    }).then(function (response) {
      {logInfo("Fetched Customers Data", "getCustomerdata");}
      var resp = response.data;
      setCustomer(resp);
    })
    .catch((error) => {
      logError("Failed to fetch CustomersData", "getCustomerdata");
    });
  };

  useEffect(() => {
    getCustomerdata();
  }, []);
  const getCountries = () => {
    axios
      .get(baseUrl + `/revenuemetricsms/headCountAndTrend/getCountries`)
      .then((Response) => {{}
      logInfo("Authorizationsuccessful_getCountries", "getCountries");
        let countries = [];
        let data = Response.data;

        data.length > 0 &&
          data.forEach((e) => {
            let countryObj = {
              label: e.countryName,
              value: e.id,
            };
            countries.push(countryObj);
          });
        const sortedcities = countries.sort(function (a, b) {
          var nameA = a.label.toUpperCase();
          var nameB = b.label.toUpperCase();
          if (nameA < nameB) {
            return -1; //nameA comes first
          }
          if (nameA > nameB) {
            return 1; // nameB comes first
          }
          return 0; // names must be equal
        });
        setCountry(sortedcities);
        setSelectedCountry(countries);
      })
      .catch((error) => Utils.Log(error, "error"));
  };

  useEffect(() => {
    let countryList = [];
    country.forEach((d) => {
      countryList.push(d.value);
    });
    setSearchdata((prevVal) => ({
      ...prevVal,
      ["countries"]: countryList.toString(),
    }));
  }, [country]);
  const getCadre = async () => {
    const resp = await axios({
      url: baseUrl + `/revenuemetricsms/headCountAndTrend/getCadres`,
    });

    let cadres = [];
    let data = resp.data;
    data.push({ id: "UA", cadre: "UnAssigned" });
    data.length > 0 &&
      data.forEach((e) => {
        let cadresObj = {
          label: e.cadre,
          value: e.id,
        };

        cadres.push(cadresObj);
        setCadres(cadres);
        setSelectedCadres(cadres.filter((ele) => ele.value != 0));
        let filteredCadresData = [];
        cadres.forEach((data) => {
          if (data.value != 0) {
            filteredCadresData.push(data.value);
          }
        });

        setSearchdata((prevVal) => ({
          ...prevVal,
          ["cadres"]: filteredCadresData.toString(),
        }));
      });
  };
  const getMeasure = async () => {
    const resp = await axios({
      url: baseUrl + `/revenuemetricsms/headCountAndTrend/getMeasures`,
    });

    let measures = [];
    let data = resp.data;

    data.length > 0 &&
      data.forEach((e) => {
        let measuresObj = {
          label: e.measures,
          value: e.id,
        };

        measures.push(measuresObj);

        if (selectType != "BusinessUnit") {
          measures = measures.filter(
            (item) => item.label !== "Billed Utilization"
          );
        }

        setMeasures(measures);
        setSelectedMeasures(measures.filter((ele) => ele.value != 0));
        let filteredMeasuresData = [];
        measures.forEach((data) => {
          if (data.value != 0) {
            filteredMeasuresData.push(data.value);
          }
        });
        setSearchdata((prevVal) => ({
          ...prevVal,
          ["measures"]: filteredMeasuresData.toString(),
        }));
      });
  };
  const getProjectdata = () => {
    axios({
      method: "get",
      url: baseUrl + `/ProjectMS/Audit/getProjectNameandId`,
    }).then(function (response) {
      {logInfo("Fetched Projectdata", "getProjectdata");}
      var resp = response.data;
      setProject(resp);
    })
    .catch((error) => {
      logError("Failed to fetch Projectdata", "getProjectdata");
    });
  };
  const getDeliveryPartners = () => {
    axios
      .get(baseUrl + `/administrationms/subkconversiontrend/getdeliverypartner`)
      .then((Response) => {
        {logInfo("DeliveryPartners fetched successfully", "getDeliveryPartners");}
        let deliver = [];
        let data = Response.data;
        data.length > 0 &&
          data.forEach((e) => {
            let deliverObj = {
              label: e.PersonName,
              value: e.id,
            };
            deliver.push(deliverObj);
          });
        setDelivery(deliver);
        setSelectedDelivery(deliver);
      })
      .catch((error) => {
        logError("Failed to fetch DeliveryPartners", "getDeliveryPartners");
      });
  };

  const getAccountOwner = () => {
    axios({
      method: "get",
      url: baseUrl + `/ProjectMS/Engagement/getAccountOwner`,
    }).then((res) => {
      {logInfo("Fetched AccountOwner", "getAccountOwner");}
      let custom = [];
      let data = res.data;
      data.length > 0 &&
        data.forEach((e) => {
          let dpObj = {
            label: e.Name,
            value: e.account_owner_id,
          };
          custom.push(dpObj);
        });
      custom.push({ label: "UnAssigned", account_owner_id: 999 });
      setAccountOwner(custom);
      setSelectedAccountOwner(custom);
    })
    .catch((error) => {
      logError("Failed to fetch AccountOwner", "getAccountOwner");
    });
  };

  const getEngCompany = async () => {
    const resp = await axios({
      url: baseUrl + `/revenuemetricsms/headCountAndTrend/getEngCompany`,
    });

    let EngComp = [];
    let data = resp.data;

    data.length > 0 &&
      data.forEach((e) => {
        let EngCompObj = {
          label: e.label,
          value: e.value,
        };

        EngComp.push(EngCompObj);
        let filteredEngCompany = EngComp.filter(
          (ele) => ![0, 10, 11, 12, 13].includes(ele.value)
        );
        setEngComp(filteredEngCompany);

        setSelectedEngComp(filteredEngCompany);

        let filteredEngComp = [];
        EngComp.forEach((data) => {
          if (data.value != 0) {
            filteredEngComp.push(data.value);
          }
        });
        setSearchdata((prevVal) => ({
          ...prevVal,
          ["engComp"]: -1,
        }));
      });
  };
    
  useEffect(() => {
    getDepartments();
    getCountries();
    getCadre();
    getMeasure();
    handleCsl();
    getProjectdata();
    getDeliveryPartners();
    getAccountOwner();
    getEngCompany();
  }, [selectType]);
  const HelpPDFName = "HeadCountTrend.pdf";
  const Headername = "Head Count Trend Help";

  const generateDropdownLabel = (selectedOptions, allOptions) => {
    const selectedValues = selectedOptions.map((option) => option.value);

    const allValues = allOptions.map((item) => item.value);

    if (selectedValues.length === allValues.length) {
      return "<< ALL >>";
    } else {
      return selectedOptions.map((option) => option.label).join(", ");
    }
  };

  const [isOn, setOn] = useState(false);

  const handleToggle = () => {
    setOn(!isOn);
  };
  return (
    <div>
      {validationmessage ? (
        <div className="statusMsg error">
          {" "}
          <AiFillWarning /> Please select valid values for highlighted fields
        </div>
      ) : (
        ""
      )}

      <div className="pageTitle">
        <div className="childOne"></div>
        <div className="childTwo">
          <h2>Headcount and Margins Trend</h2>
        </div>

        <div className="childThree toggleBtns">
          <button
            className="searchFilterButton btn btn-primary"
            onClick={() => {
              setVisible(!visible);

              visible
                ? setCheveronIcon(FaChevronCircleUp)
                : setCheveronIcon(FaChevronCircleDown);
            }}
          >
            Search Filters
            <span className="serchFilterText">{cheveronIcon}</span>
          </button>
          <GlobalHelp pdfname={HelpPDFName} name={Headername} />
        </div>
      </div>

      <div>
        <div className="group mb-3 customCard">
          <div className="col-md-12 collapseHeader"></div>
          <CCollapse visible={!visible}>
            <div className="group-content row">
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="Start Month">
                    Start Month{" "}
                    <span className="required">
                      &nbsp;<span className="required error-text">*</span>
                    </span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div className="col-6 datepicker">
                    <DatePicker
                      className="month"
                      id="month"
                      name="month"
                      maxDate={
                        new Date(
                          currentDate.getFullYear(),
                          currentDate.getMonth(),
                          0
                        )
                      }
                      selected={month}
                      dateFormat="MMM-yyyy"
                      showMonthYearPicker
                      onChange={(date) => {
                        handleMonthChange(date);
                        setIsSearched(false);
                      }}
                    />
                  </div>
                </div>
              </div>
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="duration">
                    Duration&nbsp;<span className="required error-text">*</span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div className="col-6">
                    <select
                      onChange={handleChange}
                      className="text"
                      defaultValue={3}
                      name="duration"
                      id="duration"
                      ref={(ele) => {
                        ref.current[0] = ele;
                      }}
                    >
                      <option value="">&lt;&lt; Please Select &gt;&gt;</option>
                      {[...Array(Math.min(diffMonths, 12))].map((_, index) => {
                        const optionValue = index + 1;
                        return (
                          <option key={optionValue} value={optionValue}>
                            {optionValue}
                          </option>
                        );
                      })}
                    </select>
                  </div>
                </div>
              </div>
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="searchType">
                    Search Type&nbsp;
                    <span className="required error-text">*</span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div className="col-6">
                    <select
                      onChange={handleChange}
                      className="text"
                      defaultValue="BusinessUnit"
                      name="searchType"
                      id="searchType"
                      ref={(ele) => {
                        ref.current[1] = ele;
                      }}
                    >
                      <option value="">&lt;&lt; Please Select &gt;&gt;</option>
                      <option value="BusinessUnit">Business Unit</option>
                      <option value="CSL">CSL</option>
                      <option value="Customer">Customer</option>
                      <option value="DP">Delivery Partner</option>
                      <option value="Project">Project</option>
                      <option value="AccountOwner">Account Owner</option>
                      <option value="engComp">Eng Company</option>
                    </select>
                  </div>
                </div>
              </div>

              {selectType == "BusinessUnit" ? (
                <div className=" col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5" htmlFor="Business Unit">
                      Business Unit{" "}
                      <span className="required">
                        &nbsp;<span className="required error-text">*</span>
                      </span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div
                      className="col-6 multiselect"
                      ref={(ele) => {
                        ref.current[2] = ele;
                      }}
                    >
                      <MultiSelect
                        ArrowRenderer={ArrowRenderer}
                        id="busUnits"
                        options={departments.sort((a, b) =>
                          a.label.localeCompare(b.label)
                        )}
                        hasSelectAll={true}
                        isLoading={false}
                        shouldToggleOnHover={false}
                        disableSearch={false}
                        value={selectedDepartments}
                        disabled={false}
                        onChange={(s) => {
                          setSelectedDepartments(s);
                          let filteredValues = [];
                          s.forEach((d) => {
                            filteredValues.push(d.value);
                          });

                          setSearchdata((prevVal) => ({
                            ...prevVal,
                            ["busUnits"]: filteredValues.toString(),
                          }));
                        }}
                        valueRenderer={generateDropdownLabel}
                      />
                    </div>
                  </div>
                </div>
              ) : selectType == "Customer" ? (
                <div className=" col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5" htmlFor="Customer">
                      Customer&nbsp;
                      <span className="error-text">*</span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div className="col-6">
                      <select
                        className="cancel text"
                        name="Customer"
                        id="Customer"
                        onChange={handleChange}
                        defaultValue={0}
                      >
                        <option value={-1}> &lt;&lt;All&gt;&gt;</option>
                        <option value={0}>Active Customers</option>
                        <option value={1}>Select</option>
                      </select>
                    </div>
                  </div>
                </div>
              ) : selectType == "Project" ? (
                <div className="col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5" htmlFor="email-input">
                      Project&nbsp;<span className="error-text ml-1">*</span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div className="col-6">
                      <div
                        className="autoComplete-container react  reactsearchautocomplete"
                        ref={(ele) => {
                          ref.current[2] = ele;
                        }}
                      >
                        <ReactSearchAutocomplete
                          items={project}
                          type="Text"
                          name="prjId"
                          id="prjId"
                          className="error AutoComplete"
                          onSelect={(e) => {
                            setSearchdata((prevProps) => ({
                              ...prevProps,
                              prjId: e.id,
                            }));
                          }}
                          showIcon={false}
                          placeholder="Type minimum 3 characters"
                          maxResults={1000}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              ) : selectType == "CSL" ? (
                <div className=" col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5" htmlFor="CSL">
                      CSL&nbsp;
                      <span className="error-text">*</span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div
                      className="col-6 multiselect"
                      ref={(ele) => {
                        ref.current[2] = ele;
                      }}
                    >
                      <MultiSelect
                        ArrowRenderer={ArrowRenderer}
                        id="cslRes"
                        name="cslRes"
                        options={csl.sort((a, b) =>
                          a.label.localeCompare(b.label)
                        )}
                        hasSelectAll={true}
                        value={selectedCsl}
                        shouldToggleOnHover={false}
                        valueRenderer={generateDropdownLabel}
                        disableSearch={false}
                        selected={selectedCsl}
                        disabled={false}
                        onChange={(e) => {
                          setSelectedCsl(e);
                          let filteredCsl = [];
                          e.forEach((d) => {
                            filteredCsl.push(d.value);
                          });
                          setSearchdata((prevVal) => ({
                            ...prevVal,
                            ["cslRes"]: filteredCsl.toString(),
                          }));
                        }}
                      />
                    </div>
                  </div>
                </div>
              ) : selectType == "DP" ? (
                <div className=" col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5" htmlFor="businessunit">
                      Delivery Partner&nbsp;
                      <span className="error-text">*</span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div
                      className="col-6 multiselect"
                      ref={(ele) => {
                        ref.current[2] = ele;
                      }}
                    >
                      <MultiSelect
                        ArrowRenderer={ArrowRenderer}
                        id="dpRes"
                        options={delivery.sort((a, b) =>
                          a.label.localeCompare(b.label)
                        )}
                        hasSelectAll={true}
                        isLoading={false}
                        valueRenderer={generateDropdownLabel}
                        value={selectedDelivery}
                        disabled={false}
                        disableSearch={false}
                        shouldToggleOnHover={false}
                        onChange={(e) => {
                          setSelectedDelivery(e);
                          let filteredDelivery = [];
                          e.forEach((d) => {
                            filteredDelivery.push(d.value);
                          });
                          setSearchdata((prevVal) => ({
                            ...prevVal,
                            ["dpRes"]: filteredDelivery.toString(),
                          }));
                        }}
                      />
                    </div>
                  </div>
                </div>
              ) : selectType == "AccountOwner" ? (
                <div className=" col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5 ellipsis" htmlFor="Dropdown2">
                      Account Owner&nbsp;
                      <span className="error-text">*</span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div
                      className="col-6 multiselect"
                      ref={(ele) => {
                        ref.current[3] = ele;
                      }}
                    >
                      <MultiSelect
                        ArrowRenderer={ArrowRenderer}
                        id="accOwnerRes"
                        name="accOwnerRes"
                        options={accountOwner}
                        hasSelectAll={true}
                        value={selectedAccountOwner}
                        valueRenderer={generateDropdownLabel}
                        shouldToggleOnHover={false}
                        disableSearch={false}
                        disabled={false}
                        onChange={(e) => {
                          setSelectedAccountOwner(e);
                          let filteredCountry = [];
                          e.forEach((d) => {
                            filteredCountry.push(d.value);
                          });

                          setSearchdata((prevVal) => ({
                            ...prevVal,
                            ["accOwnerRes"]: filteredCountry.toString(),
                          }));
                        }}
                      />
                    </div>
                  </div>
                </div>
              ) : selectType == "engComp" ? (
                <div className=" col-md-3 mb-2">
                  <div className="form-group row">
                    <label className="col-5 ellipsis" htmlFor="Dropdown2">
                      Eng.Company&nbsp;
                      <span className="error-text">*</span>
                    </label>
                    <span className="col-1 p-0">:</span>
                    <div
                      className="col-6 multiselect"
                      ref={(ele) => {
                        ref.current[3] = ele;
                      }}
                    >
                      <MultiSelect
                        ArrowRenderer={ArrowRenderer}
                        id="engComp"
                        name="engComp"
                        options={engComp}
                        hasSelectAll={true}
                        value={selectedEngComp}
                        valueRenderer={generateDropdownLabel}
                        shouldToggleOnHover={false}
                        disableSearch={false}
                        disabled={false}
                        onChange={(e) => {
                          setSelectedEngComp(e);
                          let filteredCountry = [];
                          e.forEach((d) => {
                            filteredCountry.push(d.value);
                          });

                          setSearchdata((prevVal) => ({
                            ...prevVal,
                            ["engComp"]: filteredCountry.toString(),
                          }));
                        }}
                      />
                    </div>
                  </div>
                </div>
              ) : (
                ""
              )}

              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="Res. Type">
                    Res. Type&nbsp;
                    <span className="required error-text">*</span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div className="col-6">
                    <select onChange={handleChange} name="resTyp" id="resTyp">
                      <option value={-1}>&lt;&lt; All &gt;&gt;</option>
                      <option value="subK">Contractor</option>
                      <option value="fte">Employee</option>
                    </select>
                  </div>
                </div>
              </div>
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="Res.Location">
                    Res.Location{" "}
                    <span className="required">
                      &nbsp;<span className="required error-text">*</span>
                    </span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div
                    className="col-6 multiselect"
                    ref={(ele) => {
                      ref.current[selectType == "" ? 3 : 4] = ele;
                    }}
                  >
                    <MultiSelect
                      ArrowRenderer={ArrowRenderer}
                      id="countries"
                      name="countries"
                      options={country}
                      hasSelectAll={true}
                      value={selectedCountry}
                      disabled={false}
                      onChange={(e) => {
                        setSelectedCountry(e);
                        let filteredCountry = [];
                        e.forEach((d) => {
                          filteredCountry.push(d.value);
                        });
                        setSearchdata((prevVal) => ({
                          ...prevVal,
                          ["countries"]: filteredCountry.toString(),
                        }));
                      }}
                      valueRenderer={generateDropdownLabel}
                    />
                  </div>
                </div>
              </div>
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="cadres">
                    Cadres{" "}
                    <span className="required">
                      &nbsp;<span className="required error-text">*</span>
                    </span>
                  </label>
                  <span className="col-1 p-0 p-0">:</span>
                  <div
                    className="col-6 multiselect"
                    ref={(ele) => {
                      ref.current[selectType == "" ? 4 : 5] = ele;
                    }}
                  >
                    <MultiSelect
                      ArrowRenderer={ArrowRenderer}
                      id="cadres"
                      options={cadres}
                      hasSelectAll={true}
                      isLoading={false}
                      shouldToggleOnHover={false}
                      disableSearch={false}
                      value={selectedCadres}
                      disabled={false}
                      onChange={(s) => {
                        setSelectedCadres(s);
                        let filteredValues = [];
                        s.forEach((d) => {
                          filteredValues.push(d.value);
                        });

                        setSearchdata((prevVal) => ({
                          ...prevVal,
                          ["cadres"]: filteredValues.toString(),
                        }));
                      }}
                      valueRenderer={generateDropdownLabel}
                    />
                  </div>
                </div>
              </div>
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="measures">
                    Measures{" "}
                    <span className="required">
                      &nbsp;<span className="required error-text">*</span>
                    </span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div
                    className="col-6 multiselect"
                    ref={(ele) => {
                      ref.current[selectType == "" ? 5 : 6] = ele;
                    }}
                  >
                    <MultiSelect
                      ArrowRenderer={ArrowRenderer}
                      id="measures"
                      name="measures"
                      options={measures.sort((a, b) =>
                        a.label.localeCompare(b.label)
                      )}
                      hasSelectAll={true}
                      isLoading={false}
                      shouldToggleOnHover={false}
                      disableSearch={false}
                      value={selectedMeasures}
                      disabled={false}
                      onChange={(s) => {
                        setSelectedMeasures(s);
                        let filteredValues = [];
                        s.forEach((d) => {
                          filteredValues.push(d.value);
                        });

                        setSearchdata((prevVal) => ({
                          ...prevVal,
                          ["measures"]: filteredValues.toString(),
                        }));
                      }}
                      valueRenderer={generateDropdownLabel}
                    />
                  </div>
                </div>
              </div>
              <div className=" col-md-3 mb-2">
                <div className="form-group row">
                  <label className="col-5" htmlFor="sortBy">
                    Sort By&nbsp;<span className="required error-text">*</span>
                  </label>
                  <span className="col-1 p-0">:</span>
                  <div className="col-6">
                    <select onChange={handleChange} name="sortBy" id="sortBy">
                      <option value="-1">Default</option>
                      {selectedMeasures.map((Item) => (
                        <option value={Item.value}> {Item.label}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
              {selectType == "BusinessUnit" ? (
                <>
                  {selectedMeasures.some((item) => item.value === 1339) && (
                    <div className="col-md-3 mb-2">
                      <div className="form-group row">
                        <label className="col-5">Consider Overhead Cost</label>
                        <span className="col-1 p-0">:</span>
                        <span className="col-6 align-self-center">
                          <input
                            type="checkbox"
                            onChange={(e) => {
                              setSearchdata({
                                ...searchdata,
                                overheadChecked: e.target.checked,
                              });
                            }}
                            style={{ marginTop: "5px" }} // add margin-top here
                          ></input>
                        </span>
                      </div>
                    </div>
                  )}
                </>
              ) : (
                ""
              )}

              <div className="col-md-12 btn-container center my-3">
                <button
                  type="submit"
                  className="btn btn-primary"
                  onClick={() => {
                    handleSearch();
                  }}
                >
                  <FaSearch /> Search{" "}
                </button>
              </div>
            </div>
          </CCollapse>
        </div>
      </div>

      {loader ? <Loader handleAbort={handleAbort} /> : ""}

      <SelectCustDialogBox
        flag={1}
        visible={custVisible}
        setVisible={setCustVisible}
        setSelectedItems={setSelectedItems}
        selectedItems={selectedItems}
      />
      {toggleButton && (
        <div className="mrTableContainer">
          <div className="tableViewSwitch HMT-pagination-container">
            <span className="scrollTxt">Scroll</span>
            <div
              className={`switch-container ${isOn ? "on" : "off"}`}
              onClick={handleToggle}
            >
              <div className="switch-slider"></div>
            </div>
            <span className="paginationTxt">Pagination</span>
          </div>
        </div>
      )}
      {displayTableState && (
        <HeadCountTableComponent
          tableData={tableData}
          column={column}
          IsSearched={IsSearched}
          searchdata={searchdata}
          month={month}
          selectType={selectType}
          loggedUserId={loggedUserId}
          projects={project}
          customer={customer}
          selectedCust={selectedCust}
          laterMeasureLabel={laterMeasureLabel}
          isOn={isOn}
          visible={visible}
          reportRunId={reportRunId}
          accountOwner={accountOwner}
          engComp={engComp}
          setAccountOwner={setAccountOwner}
          selectedAccountOwner={selectedAccountOwner}
        />
      )}
    </div>
  );
}

export default HeadCount;
