function HydrogenCalc() {
  Chart.register(ChartDataLabels);
  this.$ = jQuery;
  this.chart;
  this.xlsx = XLSX;
  this.sheets = {};
  this.tabs = ["#Dashboard", "#Assumptions", "#BaseCase", "#SMR", "#ATR"];
  this.init();
}

function AmmoniaCalc() {
  this.$ = jQuery;
  this.chart2;
  this.xlsx = XLSX;
  this.sheets = {};
  this.tabs = ["#Dashboard2", "#Assumptions2", "#KBRPurifierReferenceCase", "#KBRPurifierList2"];
  this.init();
}

function GreenHydrogenCalc() {
  this.$ = jQuery;
  this.chart3;
  this.chart4;
  this.xlsx = XLSX;
  this.sheets = {};
  this.tabs = ["#Dashboard3", "#Assumptions3", "#H2PEM", "#H2AEL", "#H2SOEC", "#NH3PEM", "#NH3AEL", "#NH3SOEC"];
  this.init();
}

function onHydrogen() {
  document.querySelector("#hydrogenCalc").scrollIntoView(true);
}

function onAmmonia() {
  document.querySelector("#ammoniaCalc").scrollIntoView(true);
}

function onGreen() {
  document.querySelector("#greenCalc").scrollIntoView(true);
}

HydrogenCalc.fn = HydrogenCalc.prototype;
AmmoniaCalc.fn = AmmoniaCalc.prototype;
GreenHydrogenCalc.fn = GreenHydrogenCalc.prototype;

let self = null;
let self2 = null;
let self3 = null;

let chartOpts;
let chartOpts2;
let chartOpts3;
let chartOpts4;

function hiddenLoader() {
  setTimeout(function() {
    getElById("loaderMain").style = "display: none";
    getElById("body").style.overflow = "auto"
  }, 2000);
}

function formulaParse(formula, sheetNames) {
  let cleanFormula = formula.split("$").join("");
  //Removing spaces in sheet names
  sheetNames.forEach(sheetName => {
    cleanFormula = cleanFormula.split(sheetName).join(`#${sheetName.split(" ").join("")}`);
  });
  cleanFormula = cleanFormula.split("'").join("");
  cleanFormula = cleanFormula.split(">").join(" > ");
  cleanFormula = cleanFormula.split("<").join(" < ");
  return cleanFormula;
}

function splitName(name) {
  return name.split(" ").join("");
}

function setAttrValue(element, value) {
  element.value = value
  // element.setAttribute("value", value);
}

function getElById(id) {
  return document.getElementById(id);
}

function buildValue(element) {
  const value = element.value;
  return value?.replace(/^[-+]?[0-9]*[.,]?[0-9]+$/g, "").replace("$", "").replace(",", "");
}

let dashboard = null;
HydrogenCalc.fn.init = async function() {
  getElById("body").style.overflow = "hidden";
  self = this;

  const f = await fetch("https://docs.google.com/spreadsheets/d/1fdP3vMapCDwfEC7fRxkBpzD5ODnhXEwH79U8a7RU1Ic/export?format=xlsx");

  const a = await f.arrayBuffer();
  const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
  if (wb) {
    hiddenLoader();
  }
  const taxCreditEl = getElById("taxCredit");
  const checkboxDataValueTaxCr = wb.Sheets["Dashboard"]["G15"].v;
  if ((checkboxDataValueTaxCr == "Yes" && !taxCreditEl.checked) || (checkboxDataValueTaxCr && taxCreditEl.checked)) {
    taxCreditEl.click();
  }
  const electricityExportEl = getElById("electricityExport");
  const checkboxDataValueEl = wb.Sheets["Dashboard"]["G8"].v;
  if ((checkboxDataValueEl == "Yes" && !electricityExportEl.checked) || (checkboxDataValueEl && electricityExportEl.checked)) {
    electricityExportEl.click();
  }

  const gasVisEl = getElById("gasVis");
  const checkboxDataValueGas = wb.Sheets["Dashboard"]["G11"].v;
  setAttrValue(gasVisEl, checkboxDataValueGas);


  const electricityVisEl = getElById("electricityVis");
  const checkboxDataValueElectricity = wb.Sheets["Dashboard"]["G12"].v;
  setAttrValue(electricityVisEl, checkboxDataValueElectricity);


  const carbonVisEl = getElById("carbonVis");
  const checkboxDataValueCarbon = wb.Sheets["Dashboard"]["G13"].v;
  setAttrValue(carbonVisEl, checkboxDataValueCarbon);


  const carbonPriceVisEl = getElById("carbonPriceVis");
  const checkboxDataValueCarbonPrice = wb.Sheets["Dashboard"]["G14"].v;
  setAttrValue(carbonPriceVisEl, checkboxDataValueCarbonPrice);


  const hydrogenData = {};
  Object.keys(wb.Sheets).forEach(name => {
    Object.keys(wb.Sheets[name]).forEach(cell => {
      if (!hydrogenData[splitName(name)]) {
        hydrogenData[splitName(name)] = {};
      }
      // We only need cells
      if (cell[0] === "!") {
        return;
      }
      if (wb.Sheets[name][cell].f) {
        hydrogenData[splitName(name)][cell] = {
          format: wb.Sheets[name][cell].z || "",
          formula: formulaParse(wb.Sheets[name][cell].f, wb.SheetNames),
          value: wb.Sheets[name][cell].v
        };
      } else {
        hydrogenData[splitName(name)][cell] = {
          format: wb.Sheets[name][cell].z || "",
          value: wb.Sheets[name][cell].v
        };
      }
      if (hydrogenData[splitName(name)][cell].format === "General") {
        hydrogenData[splitName(name)][cell].format = "";
      }
    });
  });


  $(self.tabs.join(",")).calx({
    data: hydrogenData,
    onAfterCalculate: function() {
      if (self.chart) {
        self.chart.data.datasets[0].data = [];
        self.chart.update();
      }
    }
  });

  self.tabs.map(function(tab) {
    self.sheets[tab.replace("#", "")] = self.$(tab).calx("getSheet");
  });

  setTimeout(function() {
    self.drawChart();
  }, 2000);
};

HydrogenCalc.fn.getDefaultChartOpts = function() {
  return {
    type: "bar",
    data: {
      labels: ["Reference Case", "SMR +90% CCS", "ATR + GHR"],
      datasets: []
    },
    options: {
      plugins: {
        datalabels: {
          formatter: function(value, context) {
            return "";
          }
        }
      },
      scales: {
        x: {
          stacked: true
        },
        y: {
          stacked: true,
          title: {
            display: true,
            text: "Levelized Cost of Hydrogen (USD/kg)"
          }
        }
      }
    },
    plugins: [ChartDataLabels, {
      id: "topLabels",
      afterDatasetsDraw(chart, args, pluginOpt) {
        const { ctx, scales: { x, y } } = chart;
        chart.data.datasets[0].data.forEach((datapoint, index) => {
          const totalArray = [buildValue(getElById("total1")), buildValue(getElById("total2")), buildValue(getElById("total3"))];
          totalArray.forEach((data, index) => {
            ctx.font = "bold";
            ctx.fillStyle = "#808080";
            ctx.textAlign = "center";
            ctx.fillText(data, x.getPixelForValue(index), chart.getDatasetMeta(7).data[index].y - 5);
          });
        });
      }
    }]
  };
};


HydrogenCalc.fn.getVal = function(sheet, addr) {
  return this.sheets[sheet ?? "SMR"].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
};

HydrogenCalc.fn.drawChart = function() {
  self = this;
  dashboard = self.sheets["Dashboard"];
  chartOpts = self.getDefaultChartOpts();
  chartOpts.data.datasets = [{
    label: "Carbon Credit",
    data: [
      self.getVal("BaseCase", "K128"),
      self.getVal("BaseCase", "L128"),
      self.getVal("BaseCase", "M128")
    ],
    backgroundColor: [
      "rgba(25, 9, 232, 0.2)"
    ],
    borderColor: [
      "rgba(25, 9, 232, 1)"
    ],
    borderWidth: 1
  },
    {
      label: "Carbon Tax",
      data: [
        self.getVal("BaseCase", "K127"),
        self.getVal("BaseCase", "L127"),
        self.getVal("BaseCase", "M127")
      ],
      backgroundColor: [
        "rgba(173,171,68,0.2)"
      ],
      borderColor: [
        "rgb(245,237,3)"
      ],
      borderWidth: 1
    },
    {
      label: "CO2 T&S",
      data: [
        self.getVal("BaseCase", "K126"),
        self.getVal("BaseCase", "L126"),
        self.getVal("BaseCase", "M126")
      ],
      backgroundColor: [
        "rgba(136,98,65,0.2)"
      ],
      borderColor: [
        "rgb(241,133,3)"
      ],
      borderWidth: 1
    },
    {
      label: "Water",
      data: [
        self.getVal("BaseCase", "K125"),
        self.getVal("BaseCase", "L125"),
        self.getVal("BaseCase", "M125")
      ],
      backgroundColor: [
        "rgba(66,93,164,0.2)"
      ],
      borderColor: [
        "rgb(14,34,243)"
      ],
      borderWidth: 1
    },
    {
      label: "Electricity",
      data: [
        self.getVal("BaseCase", "K124"),
        self.getVal("BaseCase", "L124"),
        self.getVal("BaseCase", "M124")
      ],
      backgroundColor: [
        "rgba(68,134,60,0.2)"
      ],
      borderColor: [
        "rgb(14,227,7)"
      ],
      borderWidth: 1
    },
    {
      label: "Natural Gas",
      data: [
        self.getVal("BaseCase", "K123"),
        self.getVal("BaseCase", "L123"),
        self.getVal("BaseCase", "M123")
      ],
      backgroundColor: [
        "rgb(167,186,227)"
      ],
      borderColor: [
        "rgb(7,77,227)"
      ],
      borderWidth: 1
    },
    {
      label: "Fixed OPEX",
      data: [
        self.getVal("BaseCase", "K122"),
        self.getVal("BaseCase", "L122"),
        self.getVal("BaseCase", "M122")
      ],
      backgroundColor: [
        "rgba(61,65,63,0.2)"
      ],
      borderColor: [
        "rgb(44,43,34)"
      ],
      borderWidth: 1
    },
    {
      label: "CAPEX",
      data: [
        self.getVal("BaseCase", "K121"),
        self.getVal("BaseCase", "L121"),
        self.getVal("BaseCase", "M121")
      ],
      backgroundColor: [
        "rgba(238,96,96,0.2)"
      ],
      borderColor: [
        "rgb(227,5,43)"
      ],
      borderWidth: 1
    },
    {
      label: "Power Export",
      data: [
        self.getVal("BaseCase", "K120"),
        self.getVal("BaseCase", "L120"),
        self.getVal("BaseCase", "M120")
      ],
      backgroundColor: "rgba(75,97,182,0.2)",
      borderColor: [
        "rgb(72,15,217)"
      ],
      borderWidth: 1
    },
    {
      label: "",
      data: [0, 0, 0],
      backgroundColor: [
        "rgba(68,134,60,0)"
      ],
      borderColor: [
        "rgba(14,227,7,0)"
      ],
      borderWidth: 1
    }
  ];

  this.chart = new Chart($("#chart_container")[0], chartOpts);
};


let yesNoInputs = document.querySelectorAll('.yesNoInputs');
yesNoInputs.forEach((el)=>{
  el.addEventListener('click', function(){
    if(el.checked){
      setAttrValue(el, "Yes")
    } else {
      setAttrValue(el, "No")
    }
  })
})

let blueHydrogenInputEl = document.querySelectorAll(".blueHydrogenInput");
blueHydrogenInputEl.forEach((el)=> {
  el.addEventListener('change', buildNewCanvasHydrogenCalc)
  el.addEventListener("keypress", function(e) {
    if (e.which == 13) {
      el.blur();
    }
  })
})

function calculateHydrogenCalc() {
  $("#Assumptions").calx("getSheet").calculate();
  $("#BaseCase").calx("getSheet").calculate();
  $("#SMR").calx("getSheet").calculate();
  $("#ATR").calx("getSheet").calculate();
  $("#Dashboard").calx("getSheet").calculate();
  buildNewCanvasHydrogenCalc();
}

getElById("gasVis").addEventListener("change", function(){
  self.sheets["Dashboard"].getCell("G11").setValue(getElById("gasVis").value * 1);
  getElById("gas").blur();
  calculateHydrogenCalc();
})

getElById("electricityVis").addEventListener("change", function() {
  self.sheets["Dashboard"].getCell("G12").setValue(getElById("electricityVis").value * 1);
  getElById("electricity").blur();
  calculateHydrogenCalc();
})

getElById("carbonVis").addEventListener("change", function() {
  self.sheets["Dashboard"].getCell("G13").setValue(getElById("carbonVis").value * 1);
  getElById("carbon").blur();
  calculateHydrogenCalc();
})

getElById("carbonPriceVis").addEventListener("change", function() {
  self.sheets["Dashboard"].getCell("G14").setValue(getElById("carbonPriceVis").value * 1);
  getElById("carbonPrice").blur();
  calculateHydrogenCalc();
})

function buildNewCanvasHydrogenCalc() {
  if (chartOpts) {
    setTimeout(function() {
      chartOpts.data.datasets = [{
        label: "Carbon Credit",
        data: [
          buildValue(getElById("ccr1")),
          buildValue(getElById("ccr2")),
          buildValue(getElById("ccr2"))
        ],
        backgroundColor: [
          "rgba(25, 9, 232, 0.2)"
        ],
        borderColor: [
          "rgba(25, 9, 232, 1)"
        ],
        borderWidth: 1
      }, {
        label: "Carbon Tax",
        data: [
          buildValue(getElById("ct1")),
          buildValue(getElById("ct2")),
          buildValue(getElById("ct3"))],
        backgroundColor: [
          "rgba(173,171,68,0.2)"
        ],
        borderColor: [
          "rgb(245,237,3)"
        ],
        borderWidth: 1
      }, {
        label: "CO2 T&S",
        data: [
          buildValue(getElById("co1")),
          buildValue(getElById("co2")),
          buildValue(getElById("co3"))
        ],
        backgroundColor: [
          "rgba(136,98,65,0.2)"
        ],
        borderColor: [
          "rgb(241,133,3)"
        ],
        borderWidth: 1
      }, {
        label: "Water",
        data: [
          buildValue(getElById("wat1")),
          buildValue(getElById("wat2")),
          buildValue(getElById("wat3"))
        ],
        backgroundColor: [
          "rgba(66,93,164,0.2)"
        ],
        borderColor: [
          "rgb(14,34,243)"
        ],
        borderWidth: 1
      }, {
        label: "Electricity",
        data: [
          buildValue(getElById("el1")),
          buildValue(getElById("el2")),
          buildValue(getElById("el3"))
        ],
        backgroundColor: [
          "rgba(68,134,60,0.2)"
        ],
        borderColor: [
          "rgb(14,227,7)"
        ],
        borderWidth: 1
      }, {
        label: "Natural Gas",
        data: [
          buildValue(getElById("natGas1")),
          buildValue(getElById("natGas2")),
          buildValue(getElById("natGas3"))
        ],
        backgroundColor: [
          "rgb(167,186,227)"
        ],
        borderColor: [
          "rgb(7,77,227)"
        ],
        borderWidth: 1
      }, {
        label: "Fixed OPEX",
        data: [
          buildValue(getElById("fo1")),
          buildValue(getElById("fo2")),
          buildValue(getElById("fo3"))
        ],
        backgroundColor: [
          "rgba(61,65,63,0.2)"
        ],
        borderColor: [
          "rgb(44,43,34)"
        ],
        borderWidth: 1
      }, {
        label: "CAPEX",
        data: [
          buildValue(getElById("cap1")),
          buildValue(getElById("cap2")),
          buildValue(getElById("cap3"))
        ],
        backgroundColor: [
          "rgba(238,96,96,0.2)"
        ],
        borderColor: [
          "rgb(227,5,43)"
        ],
        borderWidth: 1
      }, {
        label: "Power Export",
        data: [
          buildValue(getElById("pe1")),
          buildValue(getElById("pe2")),
          buildValue(getElById("pe3"))
        ],
        backgroundColor: "rgba(75,97,182,0.2)",
        borderColor: [
          "rgb(72,15,217)"
        ],
        borderWidth: 1
      },
        {
          label: "",
          data: [0, 0, 0],
          backgroundColor: [
            "rgba(68,134,60,0)"
          ],
          borderColor: [
            "rgba(14,227,7,0)"
          ],
          borderWidth: 1
        }
      ];
      getElById("canvas_wr").innerHTML = "";
      getElById("canvas_wr").innerHTML = "<canvas id=\"chart_container\" height=\"200px\"></canvas>";
      this.chart = new Chart(getElById("chart_container"), chartOpts);
    }, 3000);
  }
}


//AMMONIA CALC ________________________________________________________________________________________________________

AmmoniaCalc.fn.init = async function() {
  self2 = this;
  const f = await fetch("https://docs.google.com/spreadsheets/d/136gX-lPlD2fMbYbxCJmtu_f762j5SfbAuUohJT_MdkU/export?format=xlsx");
  const a = await f.arrayBuffer();
  const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
  getElById("body").style.overflow = "hidden";

  if (wb) {
    hiddenLoader();
  }

  const taxCreditAmEl = getElById("taxCreditAm");
  const checkboxDataValueTaxCreditAm = wb.Sheets["Dashboard2"]["G12"].v;
  if ((checkboxDataValueTaxCreditAm == "Yes" && !taxCreditAmEl.checked) || (checkboxDataValueTaxCreditAm && taxCreditAmEl.checked)) {
    taxCreditAmEl.click();
  }

  const elExportAmEl = getElById("elExportAm");
  const checkboxDataValueElExportAm = wb.Sheets["Dashboard2"]["G13"].v;
  if ((checkboxDataValueElExportAm == "Yes" && !elExportAmEl.checked) || (checkboxDataValueElExportAm && elExportAmEl.checked)) {
    elExportAmEl.click();
  }

  const gasAmVisEl = getElById("gasAmVis");
  const checkboxDataValueGasAm = wb.Sheets["Dashboard2"]["G8"].v;
  setAttrValue(gasAmVisEl, checkboxDataValueGasAm);

  const electricityAmVisEl = getElById("electricityAmVis");
  const checkboxDataValueElectricityAm = wb.Sheets["Dashboard2"]["G9"].v;
  setAttrValue(electricityAmVisEl, checkboxDataValueElectricityAm);

  const carbonAmVisEl = getElById("carbonAmVis");
  const checkboxDataValueCarbonAm = wb.Sheets["Dashboard2"]["G10"].v;
  setAttrValue(carbonAmVisEl, checkboxDataValueCarbonAm);

  const carbonPriceAmVisEl = getElById("carbonPriceAmVis");
  const checkboxDataValueCarbonPriceAm = wb.Sheets["Dashboard2"]["G11"].v;
  setAttrValue(carbonPriceAmVisEl, checkboxDataValueCarbonPriceAm);

  const ammoniaData = {};
  Object.keys(wb.Sheets).forEach(name => {
    Object.keys(wb.Sheets[name]).forEach(cell => {
      if (!ammoniaData[splitName(name)]) {
        ammoniaData[splitName(name)] = {};
      }
      // We only need cells
      if (cell[0] === "!") {
        return;
      }
      if (wb.Sheets[name][cell].f) {
        ammoniaData[splitName(name)][cell] = {
          format: wb.Sheets[name][cell].z || "",
          formula: formulaParse(wb.Sheets[name][cell].f, wb.SheetNames),
          value: wb.Sheets[name][cell].v
        };
      } else {
        ammoniaData[splitName(name)][cell] = {
          format: wb.Sheets[name][cell].z || "",
          value: wb.Sheets[name][cell].v
        };
      }
      if (ammoniaData[splitName(name)][cell].format === "General") {
        ammoniaData[splitName(name)][cell].format = "";
      }
    });
  });


  $(self2.tabs.join(",")).calx({
    data: ammoniaData,
    onAfterCalculate: function() {
      if (self2.chart) {
        self2.chart.data.datasets[0].data = [];
        self2.chart.update();
      }
    }
  });

  self2.tabs.map(function(tab) {
    self2.sheets[tab.replace("#", "")] = self2.$(tab).calx("getSheet");
  });


  setTimeout(function() {
    self2.drawChart();
  }, 2000);
};

AmmoniaCalc.fn.getDefaultChartOpts = function() {
  return {
    type: "bar",
    data: {
      labels: ["Reference", "Blue Ammonia"],
      datasets: []
    },
    options: {
      plugins: {
        datalabels: {
          formatter: function(value, context) {
            return "";
          }
        }
      },
      scales: {
        x: {
          stacked: true
        },
        y: {
          stacked: true,
          title: {
            display: true,
            text: "Levelized Cost of Ammonia (USD/ton)"
          }
        }
      }
    },
    plugins: [ChartDataLabels, {
      id: "topLabels",
      afterDatasetsDraw(chart, args, pluginOpt) {
        const { ctx, scales: { x, y } } = chart;
        chart.data.datasets[0].data.forEach((datapoint, index) => {
          const totalArray = [buildValue(getElById("amtotal1")), buildValue(getElById("amtotal2"))];
          totalArray.forEach((data, index) => {
            ctx.font = "bold";
            ctx.fillStyle = "#808080";
            ctx.textAlign = "center";
            ctx.fillText(data, x.getPixelForValue(index), chart.getDatasetMeta(8).data[index].y - 5);
          });
        });
      }
    }]
  };
};


AmmoniaCalc.fn.getVal = function(sheet, addr) {
  return this.sheets[sheet ?? "KBRPurifierReferenceCase"].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
};

AmmoniaCalc.fn.drawChart = function() {
  self2 = this;
  chartOpts2 = self2.getDefaultChartOpts();
  chartOpts2.data.datasets = [
    {
      label: "Carbon Credit",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G127"),
        self2.getVal("KBRPurifierList2", "C127")
      ],
      backgroundColor: [
        "rgba(25, 9, 232, 0.2)"
      ],
      borderColor: [
        "rgba(25, 9, 232, 1)"
      ],
      borderWidth: 1
    },
    {
      label: "Carbon Tax",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G126"),
        self2.getVal("KBRPurifierList2", "C126")
      ],
      backgroundColor: [
        "rgba(173,171,68,0.2)"
      ],
      borderColor: [
        "rgb(245,237,3)"
      ],
      borderWidth: 1
    },
    {
      label: "CO2 T&S",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G125"),
        self2.getVal("KBRPurifierList2", "C125")
      ],
      backgroundColor: [
        "rgba(136,98,65,0.2)"
      ],
      borderColor: [
        "rgb(241,133,3)"
      ],
      borderWidth: 1
    },
    {
      label: "Water",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G124"),
        self2.getVal("KBRPurifierList2", "C124")
      ],
      backgroundColor: [
        "rgba(66,93,164,0.2)"
      ],
      borderColor: [
        "rgb(14,34,243)"
      ],
      borderWidth: 1
    },
    {
      label: "Electricity",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G123"),
        self2.getVal("KBRPurifierList2", "C123")
      ],
      backgroundColor: [
        "rgba(68,134,60,0.2)"
      ],
      borderColor: [
        "rgb(14,227,7)"
      ],
      borderWidth: 1
    },
    {
      label: "Natural Gas",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G122"),
        self2.getVal("KBRPurifierList2", "C122")
      ],
      backgroundColor: [
        "rgb(167,186,227)"
      ],
      borderColor: [
        "rgb(7,77,227)"
      ],
      borderWidth: 1
    },
    {
      label: "Fixed OPEX",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G121"),
        self2.getVal("KBRPurifierList2", "C121")
      ],
      backgroundColor: [
        "rgba(61,65,63,0.2)"
      ],
      borderColor: [
        "rgb(44,43,34)"
      ],
      borderWidth: 1
    },
    {
      label: "CAPEX",
      data: [
        self2.getVal("KBRPurifierReferenceCase", "G120"),
        self2.getVal("KBRPurifierList2", "C120")
      ],
      backgroundColor: [
        "rgba(238,96,96,0.2)"
      ],
      borderColor: [
        "rgb(227,5,43)"
      ],
      borderWidth: 1
    },
    {
      label: "",
      data: [0, 0, 0],
      backgroundColor: [
        "rgba(68,134,60,0)"
      ],
      borderColor: [
        "rgba(14,227,7,0)"
      ],
      borderWidth: 1
    }
  ];

  this.chart2 = new Chart(getElById("chart_container2"), chartOpts2);
};

function buildNewCanvasAmmoniaCalc() {
  if (chartOpts2) {
    setTimeout(function() {
      chartOpts2.data.datasets = [{
        label: "CAPEX",
        data: [
          buildValue(getElById("amcap1")),
          buildValue(getElById("amcap2"))
        ],
        backgroundColor: [
          "rgba(238,96,96,0.2)"
        ],
        borderColor: [
          "rgb(227,5,43)"
        ],
        borderWidth: 1
      }, {
        label: "Fixed OPEX",
        data: [
          buildValue(getElById("amfo1")),
          buildValue(getElById("amfo2"))
        ],
        backgroundColor: [
          "rgba(61,65,63,0.2)"
        ],
        borderColor: [
          "rgb(44,43,34)"
        ],
        borderWidth: 1
      }, {
        label: "Natural Gas",
        data: [
          buildValue(getElById("amng1")),
          buildValue(getElById("amng2"))
        ],
        backgroundColor: [
          "rgb(167,186,227)"
        ],
        borderColor: [
          "rgb(7,77,227)"
        ],
        borderWidth: 1
      }, {
        label: "Electricity",
        data: [
          buildValue(getElById("amel1")),
          buildValue(getElById("amel2"))
        ],
        backgroundColor: [
          "rgba(68,134,60,0.2)"
        ],
        borderColor: [
          "rgb(14,227,7)"
        ],
        borderWidth: 1
      }, {
        label: "Water",
        data: [
          buildValue(getElById("amwat1")),
          buildValue(getElById("amwat2"))
        ],
        backgroundColor: [
          "rgba(66,93,164,0.2)"
        ],
        borderColor: [
          "rgb(14,34,243)"
        ],
        borderWidth: 1
      }, {
        label: "CO2 T&S",
        data: [
          buildValue(getElById("amco1")),
          buildValue(getElById("amco2"))
        ],
        backgroundColor: [
          "rgba(136,98,65,0.2)"
        ],
        borderColor: [
          "rgb(241,133,3)"
        ],
        borderWidth: 1
      }, {
        label: "Carbon Tax",
        data: [
          buildValue(getElById("amct1")),
          buildValue(getElById("amct2"))
        ],
        backgroundColor: [
          "rgba(173,171,68,0.2)"
        ],
        borderColor: [
          "rgb(245,237,3)"
        ],
        borderWidth: 1
      }, {
        label: "Carbon Credit",
        data: [
          buildValue(getElById("amccr1")),
          buildValue(getElById("amccr2"))
        ],
        backgroundColor: [
          "rgba(25, 9, 232, 0.2)"
        ],
        borderColor: [
          "rgba(25, 9, 232, 1)"
        ],
        borderWidth: 1
      },
        {
          label: "",
          data: [0, 0, 0],
          backgroundColor: [
            "rgba(68,134,60,0)"
          ],
          borderColor: [
            "rgba(14,227,7,0)"
          ],
          borderWidth: 1
        }];
      $("#canvas_wr2").html(""); //remove canvas from container
      $("#canvas_wr2").html("   <canvas id=\"chart_container2\" height=\"200px\"></canvas>"); //add it back to the container
      this.chart2 = new Chart(getElById("chart_container2"), chartOpts2);
    }, 3000);

  }
}

function calculateAmmoniaCalc() {
  $("#Assumptions2").calx("getSheet").calculate();
  $("#KBRPurifierReferenceCase").calx("getSheet").calculate();
  $("#KBRPurifierList2").calx("getSheet").calculate();
  $("#Dashboard2").calx("getSheet").calculate();
  buildNewCanvasAmmoniaCalc();
}

getElById("gasAmVis").addEventListener("change", function() {
  self2.sheets["Dashboard2"].getCell("G8").setValue(getElById("gasAmVis").value * 1);
  getElById("gasAm").blur();
  calculateAmmoniaCalc();
})

getElById("electricityAmVis").addEventListener("change", function() {
  self2.sheets["Dashboard2"].getCell("G9").setValue(getElById("electricityAmVis").value * 1);
  getElById("electricityAm").blur();
  calculateAmmoniaCalc();
})

getElById("carbonAmVis").addEventListener("change", function() {
  self2.sheets["Dashboard2"].getCell("G10").setValue(getElById("carbonAmVis").value * 1);
  getElById("carbonAm").blur();
  calculateAmmoniaCalc();
})

getElById("carbonPriceAmVis").addEventListener("change", function() {
  self2.sheets["Dashboard2"].getCell("G11").setValue(getElById("carbonPriceAmVis").value * 1);
  getElById("carbonPriceAm").blur();
  calculateAmmoniaCalc();
})

let blueAmmoniaInputEl = document.querySelectorAll(".blueAmmoniaInput");
blueAmmoniaInputEl.forEach((el)=> {
  el.addEventListener('change', buildNewCanvasAmmoniaCalc)
  el.addEventListener("keypress", function(e) {
    if (e.which == 13) {
      el.blur();
    }
  })
})


//GREEN CALC ________________________________________________________________________________________________________


GreenHydrogenCalc.fn.init = async function() {
  self3 = this;
  const f = await fetch("https://docs.google.com/spreadsheets/d/1paA9TsCHSWwIubKrXfDPVpdFjJicI2BcQsyhIp60vt0/export?format=xlsx");
  const a = await f.arrayBuffer();
  const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
  getElById("body").style.overflow = "hidden";
  if (wb) {
    hiddenLoader();
  }

  const ptcTaxCreditEl = getElById("ptcTaxCredit");
  const checkboxDataValuePtcTaxCredit = wb.Sheets["Dashboard3"]["G13"].v;
  if ((checkboxDataValuePtcTaxCredit == "Yes" && !ptcTaxCreditEl.checked) || (checkboxDataValuePtcTaxCredit == "No" && ptcTaxCreditEl.checked)) {
    ptcTaxCreditEl.click();
  }

  const capFacGrVisEl = getElById("capFacGrVis");
  const checkboxDataValueCapFac = wb.Sheets["Dashboard3"]["G9"].v;
  setAttrValue(capFacGrVisEl, checkboxDataValueCapFac);

  const electricityGrVisEl = getElById("electricityGrVis");
  const checkboxDataValueElectricityGr = wb.Sheets["Dashboard3"]["G10"].v;
  setAttrValue(electricityGrVisEl, checkboxDataValueElectricityGr);

  const greenHydrogenData = {};
  Object.keys(wb.Sheets).forEach(name => {
    Object.keys(wb.Sheets[name]).forEach(cell => {
      if (!greenHydrogenData[splitName(name)]) {
        greenHydrogenData[splitName(name)] = {};
      }
      // We only need cells
      if (cell[0] === "!") {
        return;
      }
      if (wb.Sheets[name][cell].f) {
        greenHydrogenData[splitName(name)][cell] = {
          format: wb.Sheets[name][cell].z || "",
          formula: formulaParse(wb.Sheets[name][cell].f, wb.SheetNames),
          value: wb.Sheets[name][cell].v
        };
      } else {
        greenHydrogenData[splitName(name)][cell] = {
          format: wb.Sheets[name][cell].z || "",
          value: wb.Sheets[name][cell].v
        };
      }
      if (greenHydrogenData[splitName(name)][cell].format === "General") {
        greenHydrogenData[splitName(name)][cell].format = "";
      }
    });
  });


  $(self3.tabs.join(",")).calx({
    data: greenHydrogenData,
    onAfterCalculate: function() {
      if (self3.chart) {
        self3.chart.data.datasets[0].data = [];

        self3.chart.update();
      }
    }
  });

  const checkboxDataValueElectrolyzerEfficiency = wb.Sheets["Dashboard3"]["G12"].v;
  if (checkboxDataValueElectrolyzerEfficiency == "High") {
    getElById("electrolyzerEfGrVis").setAttribute("checked", "checked")
  }

  const checkboxDataValueCapexGr = wb.Sheets["Dashboard3"]["G11"].v;
  if (checkboxDataValueCapexGr == "High") {
    getElById("capexGrVis").setAttribute("checked", "checked")
  }


  self3.tabs.map(function(tab) {
    self3.sheets[tab.replace("#", "")] = self3.$(tab).calx("getSheet");
  });

  setTimeout(function() {
    self3.drawChart();
  }, 2000);
};

GreenHydrogenCalc.fn.getDefaultChartOpts = function() {
  return {
    type: "bar",
    data: {
      labels: ["PEM", "AEL", "SOEC"],
      datasets: []
    },
    options: {
      plugins: {
        datalabels: {
          formatter: function(value, context) {
            return "";
          }
        }
      },
      scales: {
        x: {
          stacked: true
        },
        y: {
          stacked: true
        }
      }
    },
    plugins: [ChartDataLabels, {
      id: "topLabels",
      afterDatasetsDraw(chart, args, pluginOpt) {
        const { ctx, scales: { x, y } } = chart;
        let totalArray = [];
        if (chart.canvas.id == "chart_container3") {
          totalArray = [buildValue(getElById("gr1total1")), buildValue(getElById("gr1total2")), buildValue(getElById("gr1total3"))];
        } else if (chart.canvas.id == "chart_container4") {
          totalArray = [buildValue(getElById("gr2total1")), buildValue(getElById("gr2total2")), buildValue(getElById("gr2total3"))];
        }
        chart.data.datasets[0].data.forEach((datapoint, index) => {
          totalArray.forEach((data, index) => {
            ctx.font = "bold";
            ctx.fillStyle = "#808080";
            ctx.textAlign = "center";
            ctx.fillText(data, x.getPixelForValue(index), chart.getDatasetMeta(5).data[index].y - 5);
          });
        });
      }
    }]
  };
};


GreenHydrogenCalc.fn.getVal = function(sheet, addr) {
  return this.sheets[sheet ?? "NH3PEM"].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
};

let dashboard3 = null;

GreenHydrogenCalc.fn.drawChart = function() {
  self3 = this;
  dashboard3 = self3.sheets["Dashboard3"];

  chartOpts3 = self3.getDefaultChartOpts();
  chartOpts4 = self3.getDefaultChartOpts();
  chartOpts3.options.scales.y.title = {
    display: true,
    text: "Levelized Cost of Hydrogen (USD/kg)"
  };
  chartOpts4.options.scales.y.title = {
    display: true,
    text: "Levelized Cost of Ammonia (USD/tona)"
  };

  chartOpts3.data.datasets = [{
    label: "CAPEX",
    data: [
      self3.getVal("H2PEM", "H89"),
      self3.getVal("H2AEL", "C89"),
      self3.getVal("H2SOEC", "C89")
    ],
    backgroundColor: [
      "rgba(238,96,96,0.2)"
    ],
    borderColor: [
      "rgb(227,5,43)"
    ],
    borderWidth: 1
  },
    {
      label: "Tax Credit",
      data: [
        self3.getVal("H2PEM", "H88"),
        self3.getVal("H2AEL", "C88"),
        self3.getVal("H2SOEC", "C88")
      ],
      backgroundColor: [
        "rgba(173,171,68,0.2)"
      ],
      borderColor: [
        "rgb(245,237,3)"
      ],
      borderWidth: 1
    },
    {
      label: "Fixed OPEX",
      data: [
        self3.getVal("H2PEM", "H87"),
        self3.getVal("H2AEL", "C87"),
        self3.getVal("H2SOEC", "C87")
      ],
      backgroundColor: [
        "rgba(61,65,63,0.2)"
      ],
      borderColor: [
        "rgb(44,43,34)"
      ],
      borderWidth: 1
    },
    {
      label: "Water",
      data: [
        self3.getVal("H2PEM", "H86"),
        self3.getVal("H2AEL", "C86"),
        self3.getVal("H2SOEC", "C86")
      ],
      backgroundColor: [
        "rgba(66,93,164,0.2)"
      ],
      borderColor: [
        "rgb(14,34,243)"
      ],
      borderWidth: 1
    },
    {
      label: "Electricity",
      data: [
        self3.getVal("H2PEM", "H85"),
        self3.getVal("H2AEL", "C85"),
        self3.getVal("H2SOEC", "C85")
      ],
      backgroundColor: [
        "rgba(68,134,60,0.2)"
      ],
      borderColor: [
        "rgb(14,227,7)"
      ],
      borderWidth: 1
    },
    {
      label: "",
      data: [0, 0, 0],
      backgroundColor: [
        "rgba(68,134,60,0)"
      ],
      borderColor: [
        "rgba(14,227,7,0)"
      ],
      borderWidth: 1
    }
  ];

  this.chart3 = new Chart(getElById("chart_container3"), chartOpts3);

  chartOpts4.data.datasets = [{
    label: "CAPEX",
    data: [
      self3.getVal("NH3PEM", "G89"),
      self3.getVal("NH3AEL", "C89"),
      self3.getVal("NH3SOEC", "C90")
    ],
    backgroundColor: [
      "rgba(238,96,96,0.2)"
    ],
    borderColor: [
      "rgb(227,5,43)"
    ],
    borderWidth: 1
  },
    {
      label: "Tax Credit",
      data: [
        self3.getVal("NH3PEM", "G88"),
        self3.getVal("NH3AEL", "C88"),
        self3.getVal("NH3SOEC", "C89")
      ],
      backgroundColor: [
        "rgba(173,171,68,0.2)"
      ],
      borderColor: [
        "rgb(245,237,3)"
      ],
      borderWidth: 1
    },
    {
      label: "Fixed OPEX",
      data: [
        self3.getVal("NH3PEM", "G87"),
        self3.getVal("NH3AEL", "C87"),
        self3.getVal("NH3SOEC", "C88")
      ],
      backgroundColor: [
        "rgba(61,65,63,0.2)"
      ],
      borderColor: [
        "rgb(44,43,34)"
      ],
      borderWidth: 1
    },
    {
      label: "Water",
      data: [
        self3.getVal("NH3PEM", "G86"),
        self3.getVal("NH3AEL", "C86"),
        self3.getVal("NH3SOEC", "C87")
      ],
      backgroundColor: [
        "rgba(66,93,164,0.2)"
      ],
      borderColor: [
        "rgb(14,34,243)"
      ],
      borderWidth: 1
    },
    {
      label: "Electricity",
      data: [
        self3.getVal("NH3PEM", "G85"),
        self3.getVal("NH3AEL", "C85"),
        self3.getVal("NH3SOEC", "C86")
      ],
      backgroundColor: [
        "rgba(68,134,60,0.2)"
      ],
      borderColor: [
        "rgb(14,227,7)"
      ],
      borderWidth: 1
    },
    {
      label: "",
      data: [0, 0, 0],
      backgroundColor: [
        "rgba(68,134,60,0)"
      ],
      borderColor: [
        "rgba(14,227,7,0)"
      ],
      borderWidth: 1
    }
  ];

  this.chart4 = new Chart(getElById("chart_container4"), chartOpts4);
};

function buildNewCanvasGreenCalc() {
  if (chartOpts3) {
    setTimeout(function() {
      chartOpts3.data.datasets = [{
        label: "CAPEX",
        data: [
          buildValue(getElById("gr1cap1")),
          buildValue(getElById("gr1cap2")),
          buildValue(getElById("gr1cap3"))
        ],
        backgroundColor: [
          "rgba(238,96,96,0.2)"
        ],
        borderColor: [
          "rgb(227,5,43)"
        ],
        borderWidth: 1
      }, {
        label: "Fixed OPEX",
        data: [
          buildValue(getElById("gr1fo1")),
          buildValue(getElById("gr1fo2")),
          buildValue(getElById("gr1fo3"))
        ],
        backgroundColor: [
          "rgba(61,65,63,0.2)"
        ],
        borderColor: [
          "rgb(44,43,34)"
        ],
        borderWidth: 1
      }, {
        label: "Electricity",
        data: [
          buildValue(getElById("gr1el1")),
          buildValue(getElById("gr1el2")),
          buildValue(getElById("gr1el3"))
        ],
        backgroundColor: [
          "rgba(68,134,60,0.2)"
        ],
        borderColor: [
          "rgb(14,227,7)"
        ],
        borderWidth: 1
      }, {
        label: "Water",
        data: [
          buildValue(getElById("gr1wat1")),
          buildValue(getElById("gr1wat2")),
          buildValue(getElById("gr1wat3"))
        ],
        backgroundColor: [
          "rgba(66,93,164,0.2)"
        ],
        borderColor: [
          "rgb(14,34,243)"
        ],
        borderWidth: 1
      }, {
        label: "Tax Credit",
        data: [
          buildValue(getElById("gr1TaxCr1")),
          buildValue(getElById("gr1TaxCr2")),
          buildValue(getElById("gr1TaxCr3"))
        ],
        backgroundColor: [
          "rgba(25, 9, 232, 0.2)"
        ],
        borderColor: [
          "rgba(25, 9, 232, 1)"
        ],
        borderWidth: 1
      },
        {
          label: "",
          data: [0, 0, 0],
          backgroundColor: [
            "rgba(68,134,60,0)"
          ],
          borderColor: [
            "rgba(14,227,7,0)"
          ],
          borderWidth: 1
        }
      ];

      chartOpts4.data.datasets = [{
        label: "CAPEX",
        data: [
          buildValue(getElById("gr2cap1")),
          buildValue(getElById("gr2cap2")),
          buildValue(getElById("gr2cap3"))
        ],
        backgroundColor: [
          "rgba(238,96,96,0.2)"
        ],
        borderColor: [
          "rgb(227,5,43)"
        ],
        borderWidth: 1
      }, {
        label: "Fixed OPEX",
        data: [
          buildValue(getElById("gr2fo1")),
          buildValue(getElById("gr2fo2")),
          buildValue(getElById("gr2fo3"))
        ],
        backgroundColor: [
          "rgba(61,65,63,0.2)"
        ],
        borderColor: [
          "rgb(44,43,34)"
        ],
        borderWidth: 1
      }, {
        label: "Electricity",
        data: [
          buildValue(getElById("gr2el1")),
          buildValue(getElById("gr2el2")),
          buildValue(getElById("gr2el3"))
        ],
        backgroundColor: [
          "rgba(68,134,60,0.2)"
        ],
        borderColor: [
          "rgb(14,227,7)"
        ],
        borderWidth: 1
      }, {
        label: "Water",
        data: [
          buildValue(getElById("gr2wat1")),
          buildValue(getElById("gr2wat2")),
          buildValue(getElById("gr2wat3"))
        ],
        backgroundColor: [
          "rgba(66,93,164,0.2)"
        ],
        borderColor: [
          "rgb(14,34,243)"
        ],
        borderWidth: 1
      }, {
        label: "Tax Credit",
        data: [
          buildValue(getElById("gr2TaxCr1")),
          buildValue(getElById("gr2TaxCr2")),
          buildValue(getElById("gr2TaxCr3"))
        ],
        backgroundColor: [
          "rgba(25, 9, 232, 0.2)"
        ],
        borderColor: [
          "rgba(25, 9, 232, 1)"
        ],
        borderWidth: 1
      },
        {
          label: "",
          data: [0, 0, 0],
          backgroundColor: [
            "rgba(68,134,60,0)"
          ],
          borderColor: [
            "rgba(14,227,7,0)"
          ],
          borderWidth: 1
        }
      ];

      getElById("canvas_wr3").innerHTML = ""; //remove canvas from container
      getElById("canvas_wr3").innerHTML = "<canvas id=\"chart_container3\" height=\"200px\"></canvas>"; //add it back to the container
      this.chart3 = new Chart(getElById("chart_container3"), chartOpts3);

      getElById("canvas_wr4").innerHTML = ""; //remove canvas from container
      getElById("canvas_wr4").innerHTML = "<canvas id=\"chart_container4\" height=\"200px\"></canvas>"; //add it back to the container
      this.chart4 = new Chart(getElById("chart_container4"), chartOpts4);
    }, 3000);

  }
}

function calculateGreenCalc() {
  $("#H2PEM").calx("getSheet").calculate();
  $("#H2AEL").calx("getSheet").calculate();
  $("#H2SOEC").calx("getSheet").calculate();
  $("#NH3PEM").calx("getSheet").calculate();
  $("#NH3AEL").calx("getSheet").calculate();
  $("#NH3SOEC").calx("getSheet").calculate();
  $("#Assumptions3").calx("getSheet").calculate();
  $("#KBRPurifierReferenceCase").calx("getSheet").calculate();
  $("#Dashboard3").calx("getSheet").calculate();
  buildNewCanvasGreenCalc();
}

const electrolyzerEfGrVisEL = getElById("electrolyzerEfGrVis");
electrolyzerEfGrVisEL.addEventListener("change",function() {
  const select = getElById("electrolyzerEfGr");
  if (electrolyzerEfGrVisEL.checked) {
    setAttrValue(electrolyzerEfGrVisEL, "High");
    setAttrValue(select, "High")
    self3.sheets["Dashboard3"].getCell("G12").setValue("High");
  } else if (!electrolyzerEfGrVisEL.checked) {
    setAttrValue(electrolyzerEfGrVisEL, "Low");
    setAttrValue(select, "Low")
    self3.sheets["Dashboard3"].getCell("G12").setValue("Low");
  }
  $("#Dashboard3").calx("getSheet").getCell("G11").calculate();
  calculateGreenCalc();
})

const capexGrVisEL = getElById("capexGrVis");
capexGrVisEL.addEventListener("change",function() {
  const select = getElById("capexGr");
  if (capexGrVisEL.checked) {
    setAttrValue(capexGrVisEL, "High");
    setAttrValue(select, "High")
    self3.sheets["Dashboard3"].getCell("G11").setValue("High");
  } else if (!capexGrVisEL.checked) {
    setAttrValue(capexGrVisEL, "Low");
    setAttrValue(select, "Low")
    self3.sheets["Dashboard3"].getCell("G11").setValue("Low");
  }
  $("#Dashboard3").calx("getSheet").getCell("G11").calculate();
  calculateGreenCalc();
})

getElById("electricityGrVis").addEventListener("change", function() {
  self3.sheets["Dashboard3"].getCell("G10").setValue(getElById("electricityGrVis").value * 1);
  $("#Dashboard3").calx("getSheet").getCell("G10").calculate();
  getElById("electricityGr").blur();
  calculateGreenCalc();
})

getElById("capFacGrVis").addEventListener("change", function() {
  self3.sheets["Dashboard3"].getCell("G9").setValue(getElById("capFacGrVis").value * 1);
  $("#Dashboard3").calx("getSheet").getCell("G9").calculate();
  getElById("capFacGr").blur();
  calculateGreenCalc();
})

let greenHydrogenInputEl = document.querySelectorAll(".greenHydrogenInput");
greenHydrogenInputEl.forEach((el)=> {
  el.addEventListener('change', buildNewCanvasGreenCalc)
  el.addEventListener("keypress", function(e) {
    if (e.which == 13) {
      el.blur();
    }
  })
})
