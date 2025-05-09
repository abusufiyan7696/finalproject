import React, { useState, useMemo, memo, useRef } from "react";
import MaterialReactTable from "material-react-table";
import { useEffect } from "react";
import axios from "axios";
import { environment } from "../../environments/environment";
import moment from "moment";
import Loader from "../Loader/Loader";
import { FaChevronCircleRight, FaChevronCircleDown } from "react-icons/fa";
import "./HeadCountTableComponent.scss";
import useDynamicMaxHeight from "../PrimeReactTableComponent/useDynamicMaxHeight";
import Utils from "../../Utils";
import HeadCountTablePopUp from "./HeadCountTablePopUp";

function HeadCountTableComponent(props) {
  {
    /*----------------Props,useState's & useRef's---------------- */
  }
  const abortController = useRef(null);
  const {
    tableData,
    searchdata,
    month,
    selectType,
    loggedUserId,
    projects,
    customer,
    selectedCust,
    column,
    laterMeasureLabel,
    isOn,
    IsSearched,
    visible,
    reportRunId,
    accountOwner,
    engComp,
    setAccountOwner,
    selectedAccountOwner,
  } = props;

  const [data, setData] = useState([...column]);
  const [nodes, setNodes] = useState([]);
  const [departments, setDepartments] = useState([]);
  const baseUrl = environment.baseUrl;
  const [loader, setLoader] = useState(false);
  const [csl, setCsl] = useState([]);
  const [delivery, setDelivery] = useState([]);
  const [maxHeight, setMaxHeight] = useState();
  const [numberLink, setNumberLink] = useState(false);
  const [kpiName, setKpiName] = useState("");
  const [totalCount, setTotalCount] = useState("");
  const [searchVal, setSearchVal] = useState("");
  const [selectedMonth, setSelectedMonth] = useState("");
  const materialTableElement = document.getElementsByClassName(
    "materialReactExpandableTable headCountTable darkHeader"
  );

  const maxHeight1 = useDynamicMaxHeight(materialTableElement);

  {
    /*----------------Handle Abort---------------- */
  }
  useEffect(() => {
    setMaxHeight(maxHeight1);
    if (isOn) {
      setMaxHeight(maxHeight1 - 34);
    } else {
      setMaxHeight(maxHeight1 - 20);
    }
  }, [isOn, maxHeight1]);
  const handleAbort = () => {
    abortController.current && abortController.current.abort();
    setLoader(false);
  };

  useEffect(() => {
    setData(column);
  }, [column]);

  {
    /*----------------Handle Departments---------------- */
  }

  useEffect(() => {
    if (selectType == "BusinessUnit") {
      const getDepartments = async () => {
        const resp = await axios({
          url: baseUrl + `/CostMS/cost/getDepartments`,
        });
        let departments = resp.data;
        departments = departments.filter((ele) => ele.value >= 0);
        departments.push({ value: 999, label: "Non Revenue Units" });
        setDepartments(departments);
      };
      getDepartments();
    }

    {
      /*------------------Getting  DP------------------- */
    }

    if (selectType == "DP") {
      const getDeliveryPartners = () => {
        axios
          .get(
            baseUrl + `/administrationms/subkconversiontrend/getdeliverypartner`
          )
          .then((Response) => {
            let deliver = [];
            let data = Response.data;
            data.push({ id: 0, PersonName: "<<Others>>" });
            data.length > 0 &&
              data.forEach((e) => {
                let deliverObj = {
                  label: e.PersonName,
                  value: e.id,
                };
                deliver.push(deliverObj);
              });
            setDelivery(deliver);
          });
      };
      getDeliveryPartners();
    }

    {
      /*--------------------------Getting CSL---------------------- */
    }
    if (selectType == "CSL") {
      const handleCsl = () => {
        const loggedUser = loggedUserId;
        axios({
          method: "get",
          url:
            baseUrl +
            `/SalesMS/MasterController/getCslDropDownData?userId=${loggedUser}`,
        }).then((res) => {
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
        });
      };
      handleCsl();
    }
  }, []);

  {
    /*--------------------------Head Count By Cadre using Object Ids --------------------------*/
  }
  const Data = moment(month).startOf("month").format("yyyy-MM-DD");
  const HandleInsertedData = (objId, objLabel) => {
    const loaderTime = setTimeout(() => {
      setLoader(true);
    }, 2000);

    abortController.current = new AbortController();
    axios({
      method: "post",
      url:
        baseUrl +
        `/revenuemetricsms/headCountAndTrend/getHeadCountAndTrendDataByCadre`,
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
        busUnits: selectType != "BusinessUnit" ? -1 : searchdata.busUnits,
        customers:
          selectType != "Customer"
            ? -1
            : searchdata.Customer == "1"
            ? selectedCust
            : searchdata.Customer,
        prjId: selectType != "Project" ? -1 : searchdata.prjId,
        sortBy: searchdata.sortBy,
        cadres:
          searchdata.cadres == "E1,E2,E3,L,M,G,T1,UA" ? -1 : searchdata.cadres,
        cslRes: selectType != "CSL" ? -1 : searchdata.cslRes,
        dpRes: selectType != "DP" ? -1 : searchdata.dpRes,
        accOwnerRes:
          selectType != "AccountOwner"
            ? "-1"
            : (searchdata.accOwnerRes = accountOwner
                .map((item) => item.value)
                .join(",")),
        objId: objId, // change in tables according to the expandable we choose
        overheadChecked: false,
        engComp: searchdata.searchType == "engComp" ? searchdata.engComp : -1,
      },
    })
      .then((res) => {
        setLoader(false);
        clearTimeout(loaderTime);
        let respData = res.data.data;

        respData = respData.map((item) => {
          return {
            ...item,
            name: item.name + "_" + objId,
          };
        });

        setNodes((prevNodes) => {
          const grossMarginIndex = prevNodes.findIndex((node) => {
            return laterMeasureLabel == "Head Count" ||
              laterMeasureLabel == "New Hires" ||
              laterMeasureLabel == "Exits"
              ? (node.kpi.includes("SUBK") ||
                  node.kpi.includes(laterMeasureLabel)) &&
                  node.name === objLabel
              : node.kpi === laterMeasureLabel && node.name === objLabel;
          });
          if (grossMarginIndex !== -1) {
            return [
              ...prevNodes.slice(0, grossMarginIndex + 1),
              ...respData.map((data, i) => ({
                ...data,
                uniqueId: data.id,
                id: objId + i + 1,
              })),
              ...prevNodes.slice(grossMarginIndex + 1),
            ];
          } else {
            return prevNodes;
          }
        });
      })
      .catch((err) => {
        setLoader(false);
        Utils.Log(err, "Error");
      });
  };

  {
    /*----------------Adding Id,ParentId and Children---------------- */
  }

  useEffect(() => {
    let id = 1;
    for (let i = 0; i < tableData.length; i++) {
      tableData[i].id = id;
      id++;
      if (id % 100 == 16) {
        id += 985;
      }
    }

    setNodes(jsonToTree(tableData));
  }, [tableData]);

  const jsonToTree = (flatArray, options) => {
    options = {
      id: "id",
      parentId: "parentId",
      children: "subRows",
      ...options,
    };
    const dictionary = {};
    const tree = [];
    const children = options.children;
    flatArray.forEach((node) => {
      const nodeId = node[options.id];
      const nodeParentId = node[options.parentId];

      dictionary[nodeId] = {
        [children]: [],
        ...node,
        ...dictionary[nodeId],
      };
      dictionary[nodeParentId] = dictionary[nodeParentId] || { [children]: [] };
      dictionary[nodeParentId][children].push(dictionary[nodeId]);
    });

    Object.values(dictionary).forEach((obj) => {
      if (typeof obj[options.id] === "undefined") {
        tree.push(...obj[children]);
      }
    });
    return tree;
  };

  {
    /*--------------------------Dynamic Columns --------------------------*/
  }

  const dynamicColumns = [
    {
      header: (
        <div className="legendContainer ">
          {selectType === "BusinessUnit" ? (
            <div className="legend lightbrown">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="Business Unit">
                <b>Business Unit</b>
              </div>
            </div>
          ) : selectType === "DP" ? (
            <div className="legend brown">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="DP">
                <b>DP</b>
              </div>
            </div>
          ) : selectType === "CSL" ? (
            <div className="legend brown">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="CSL">
                <b>CSL</b>
              </div>
            </div>
          ) : selectType === "Customer" ? (
            <div className="legend darkgreen">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="Customer">
                <b>Customer</b>
              </div>
            </div>
          ) : selectType === "Project" ? (
            <div className="legend purple">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="Project">
                <b>Project</b>
              </div>
            </div>
          ) : selectType === "AccountOwner" ? (
            <div className="legend darkgreen">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="Customer">
                <b>Account Owner</b>
              </div>
            </div>
          ) : selectType === "engComp" ? (
            <div className="legend darkgreen">
              <div className="legendCircle"></div>
              <div className="legendTxt" title="engComp">
                <b>Eng Company</b>
              </div>
            </div>
          ) : (
            ""
          )}
        </div>
      ),
      accessorKey: "name",
      enableGrouping: false,
      GroupedCell: ({ cell }) => {
        /////////////////---------------------------------------Business Unit
        if (selectType == "BusinessUnit") {
          const department = departments.find((dep) => {
            return dep.value === cell.row.original.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});

          const handleClick = () => {
            if (department) {
              const { value, label } = department;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${value}`);
              });
              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${value}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(value, label);
              }
            }
          };

          if (department) {
            const { value, label } = department;
            const isActive = activeIcons[value];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={department.value}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="buRM" title={department.label}>
                  {department.label}
                </b>
              </span>
            );
          } else if (
            ["T1", "E1", "E2", "E3", "M", "L", "G", "Unassigned"].some(
              (substr) => cell.getValue().includes(substr)
            )
          ) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell.getValue().split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }

        /////////////////---------------------------------------Customer
        else if (selectType == "Customer") {
          const cust = customer.find((dep) => {
            return dep.id === cell.row.original.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});

          const handleClick = () => {
            if (cust) {
              const { id, fullName } = cust;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${id}`);
              });

              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [id]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${id}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [id]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(id, fullName);
              }
            }
          };

          if (cust) {
            const { id } = cust;
            const isActive = activeIcons[id];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={cust.id}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="custRM" title={cust.fullName}>
                  {cust.fullName}
                </b>
              </span>
            );
          } else if (
            ["T1", "E1", "E2", "E3", "M", "L", "G", "Unassigned"].some(
              (substr) => cell.getValue().includes(substr)
            )
          ) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell.getValue().split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }

        /////////////////---------------------------------------Project
        else if (selectType == "Project") {
          const proj = projects.find((dep) => {
            return dep.id === cell.row.original.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});

          const handleClick = () => {
            if (proj) {
              const { id, name } = proj;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${id}`);
              });

              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [id]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${id}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [id]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(id, name);
              }
            }
          };

          if (proj) {
            const { id, name } = proj;
            const isActive = activeIcons[id];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={proj.id}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="projRM" title={proj.name}>
                  {proj.name}
                </b>
              </span>
            );
          } else if (
            ["T1", "E1", "E2", "E3", "M", "L", "G", "Unassigned"].some(
              (substr) => cell.getValue().includes(substr)
            )
          ) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell.getValue().split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }

        /////////////////---------------------------------------CSL
        else if (selectType == "CSL") {
          const csls = csl.find((dep) => {
            return dep.value === cell.row.original.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});

          const handleClick = () => {
            if (csls) {
              const { value, label } = csls;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${value}`);
              });

              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${value}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(value, label);
              }
            }
          };
          if (csls) {
            const { value } = csls;
            const isActive = activeIcons[value];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={csls.value}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="resRM" title={csls.label}>
                  {csls.label}
                </b>
              </span>
            );
          } else if (
            ["T1", "E1", "E2", "E3", "M", "L", "G", "Unassigned"].some(
              (substr) => cell.getValue().includes(substr)
            )
          ) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell.getValue().split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }

        /////////////////---------------------------------------DP
        else if (selectType == "DP") {
          const dps = delivery.find((dep) => {
            return dep.value === cell.row.original.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});

          const handleClick = () => {
            if (dps) {
              const { value, label } = dps;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${value}`);
              });

              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${value}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(value, label);
              }
            }
          };
          if (dps) {
            const { value } = dps;
            const isActive = activeIcons[value];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={dps.value}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="resRM" title={dps.label}>
                  {dps.label}
                </b>
              </span>
            );
          } else if (
            ["T1", "E1", "E2", "E3", "M", "L", "G", "Unassigned"].some(
              (substr) => cell.getValue().includes(substr)
            )
          ) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell.getValue().split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }

        /////////////////---------------------------------------AO
        else if (selectType == "AccountOwner") {
          const dps = accountOwner.find((dep) => {
            return dep.value === cell.row.original.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});

          const handleClick = () => {
            if (dps) {
              const { value, label } = dps;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${value}`);
              });

              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${value}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(value, label);
              }
            }
          };
          if (dps) {
            const { value } = dps;
            const isActive = activeIcons[value];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={dps.value}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="resRM" title={dps.label}>
                  {dps.label}
                </b>
              </span>
            );
          } else if (
            ["T1", "E1", "E2", "E3", "M", "L", "G", "Unassigned"].some(
              (substr) => cell.getValue().includes(substr)
            )
          ) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell.getValue().split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }

        /////////////////---------------------------------------EC
        else if (selectType == "engComp") {
          const dps = engComp.find((dep) => {
            return dep.value === cell?.row?.original?.uniqueId;
          });

          const [activeIcons, setActiveIcons] = useState({});
          const keywords = [
            "T1_",
            "E1_",
            "E2_",
            "E3_",
            "M_",
            "L_",
            "G_",
            "Unassigned_",
          ];

          const result = keywords.some((substr) =>
            (cell.getValue() || "").startsWith(substr)
          );

          const handleClick = () => {
            if (dps) {
              const { value, label } = dps;
              const isDataPresent = nodes.some((node) => {
                return node.name.includes(`_${value}`);
              });

              if (isDataPresent) {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: false,
                }));
                setNodes((prevNodes) =>
                  prevNodes.filter((node) => !node.name.includes(`_${value}`))
                );
              } else {
                setActiveIcons((prevActiveIcons) => ({
                  ...prevActiveIcons,
                  [value]: true,
                }));
                setTimeout(() => {
                  setLoader(false);
                }, 2000);
                HandleInsertedData(value, label);
              }
            }
          };
          if (dps && !result) {
            const { value } = dps;
            const isActive = activeIcons[value];
            return (
              <span
                style={{ cursor: "pointer" }}
                value={dps.value}
                onClick={() => {
                  handleClick();
                }}
                className={"parent"}
              >
                {isActive ? <FaChevronCircleDown /> : <FaChevronCircleRight />}{" "}
                <b className="custRM" title={dps.label}>
                  {dps.label}
                </b>
              </span>
            );
          } else if (result) {
            return (
              <span
                className={"child"}
                style={{ display: "flex", justifyContent: "center" }}
              >
                <b className="resRM" title={cell.getValue()}>
                  {cell?.getValue()?.split("_")[0]}
                </b>
              </span>
            );
          } else {
            return (
              <span>
                <b>Summary</b>{" "}
              </span>
            );
          }
        }
      },
    },

    //-----------------------------------KPI-------------------------------------
    {
      header: "",
      accessorKey: `kpi`,
      Cell: ({ cell }) => (
        <span
          className={
            /(T1|E1|E2|E3|L_|M_|G_|Unassigned)/.test(cell.row.original.name)
              ? "child resRMchild"
              : selectType === "BusinessUnit"
              ? "buRMchild"
              : selectType === "Customer" ||
                selectType === "AccountOwner" ||
                selectType === "engComp"
              ? "custRMchild"
              : selectType === "Project"
              ? "projRMchild"
              : selectType === "CSL"
              ? "resRMchild"
              : selectType === "DP"
              ? "resRMchild"
              : "inherit"
          }
          style={{
            fontWeight: "bold",
          }}
          title={cell.getValue()}
        >
          {cell.getValue().includes("FTE") ||
          cell.getValue().includes("SUBK") ? (
            <>&nbsp;&nbsp; {cell.getValue()}</>
          ) : (
            cell.getValue()
          )}
        </span>
      ),
    },
    {
      header: "Total",
      accessorKey: "Total",
      Cell: ({ cell }) => {
        const { name, kpi } = cell.row.original;
        const value = Math.round(cell.getValue())?.toLocaleString("en-US");
        const child = /(T1|E1|E2|E3|L_|M_|G_|Unassigned)/.test(name);
        const isDollarKpi = [
          "Resource Direct Cost",
          "Recognized Revenue",
          "Revenue/Emp",
          "Resource Direct Cost/Emp",
          "Margin/Emp",
        ].includes(kpi);
        const isPercentageKpi =
          kpi === "Gross Margin %" || "Billed Utilization";

        return (
          <>
            {isDollarKpi && (
              <span style={{ display: "block", float: "left" }}>
                <b>$</b>
              </span>
            )}
            <span
              style={{ display: "block", float: "right" }}
              className={child ? "total" : ""}
            >
              {isDollarKpi || isPercentageKpi || child ? (
                <b>
                  {value}
                  {isPercentageKpi && value !== null && <span>%</span>}
                </b>
              ) : (
                <div>{value}</div>
              )}
            </span>
          </>
        );
      },
    },
  ];

  //------------------------------Hedrers Related to Dates---------------------------
  data.forEach((item, index) => {
    if (index > 1 && item !== "Total") {
      const dateObj = new Date(
        `${item.slice(0, 4)}-${item.slice(5, 7)}-${item.slice(8, 10)}`
      );

      const header = dateObj
        .toLocaleDateString("en-US", {
          month: "short",
          year: "numeric",
          timeZone: "UTC",
        })
        .replace(" ", "-");

      const isSpecialKpi = (kpi) =>
        [
          "Resource Direct Cost",
          "Recognized Revenue",
          "Revenue/Emp",
          "Resource Direct Cost/Emp",
          "Margin/Emp",
        ].includes(kpi);
      const getFormattedValue = (value) => value?.toLocaleString("en-US") || 0;

      dynamicColumns.push({
        header,
        accessorKey: `${item}`,
        className: "software",
        Cell: ({ cell }) => {
          const { name, kpi } = cell.row.original;
          const value = getFormattedValue(cell.getValue());
          const showDollar = isSpecialKpi(kpi);
          const showPercent = kpi === "Gross Margin %" || "Billed Utilization";
          const child = /(T1|E1|E2|E3|L_|M_|G_|Unassigned)/.test(name);

          return (
            <>
              {showDollar && (
                <span style={{ display: "block", float: "left" }}>$</span>
              )}
              <span style={{ display: "block", float: "right" }}>
                {child || showDollar || showPercent ? (
                  <>
                    {value}
                    {showPercent && <span>%</span>}
                  </>
                ) : (
                  <a
                    onClick={(e) => {
                      setSelectedMonth(header);
                      setTotalCount(value);
                      setKpiName(kpi);
                      setNumberLink(true);
                      setSearchVal(cell.row.original.uniqueId);
                    }}
                  >
                    {value}
                  </a>
                )}
              </span>
            </>
          );
        },
      });
    }
  });

  {
    /*------------------------------Return---------------------------------- */
  }
  return (
    <div>
      <div className="materialReactExpandableTable headCountTable H-and-MT-margin-table darkHeader">
        <MaterialReactTable
          columns={dynamicColumns}
          data={nodes}
          enableExpanding={(column) => column.id === "name"}
          enablePagination={isOn}
          paginateExpandedRows={true}
          enableBottomToolbar={isOn}
          enableGlobalFilter={true}
          enableDensityToggle={false}
          enableFullScreenToggle={false}
          enableHiding={false}
          enableColumnFilters={false}
          enableTopToolbar={false}
          enableColumnActions={false}
          enableSorting={false}
          filterFromLeafRows //apply filtering to all rows instead of just parent rows
          initialState={{
            showGlobalFilter: true,
            grouping: ["name"],
            expanded: true,
            density: "compact",
            columnPinning: { right: ["Total"] },
            pagination: { pageSize: 30 },
          }} //expand all rows by default
          muiTablePaginationProps={{
            labelRowsPerPage: false,
          }}
          muiTableContainerProps={{
            sx: {
              maxHeight: `${!visible ? maxHeight - 233 : maxHeight - 19}px`,
            },
          }}
          muiTableBodyProps={{
            sx: {
              "&": {
                borderRight: "1px solid #ccc",
                borderBottom: "1px solid #ccc",
              },
              "& td": {
                borderRight: "1px solid #ccc",
                height: "22px",
                padding: "0px 5px",
              },
            },
          }}
          muiTableHeadProps={{
            sx: {
              "& th": {
                borderTop: "1px solid #ccc",
                borderRight: "1px solid #ccc",
                background: "#f4f4f4 ",
                padding: "0 5px",
              },
            },
          }}
        />
      </div>
      {loader ? <Loader handleAbort={handleAbort} /> : ""}

      {numberLink && (
        <HeadCountTablePopUp
          numberLink={numberLink}
          setNumberLink={setNumberLink}
          kpiName={kpiName}
          totalCount={totalCount}
          reportRunId={reportRunId}
          setKpiName={setKpiName}
          selectedMonth={selectedMonth}
          searchVal={searchVal}
          setLoader={setLoader}
          loader={loader}
          handleAbort={handleAbort}
        />
      )}
    </div>
  );
}

export default HeadCountTableComponent;
