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
  this.tabs = ["#Dashboard3", "#Assumptions3", "#H2PEM", "#H2AEL", "#NH3PEM", "#NH3AEL"];
  this.init();
}

function onHydrogen(){
  document.querySelector("#hydrogenCalc").scrollIntoView(true)
}

function onAmmonia(){
  document.querySelector("#ammoniaCalc").scrollIntoView(true)
}

function onGreen(){
  document.querySelector("#greenCalc").scrollIntoView(true)
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

HydrogenCalc.fn.init = async function() {
  $("#body").css("overflow", "hidden");
  self = this;
  const f = await fetch("https://docs.google.com/spreadsheets/d/1fdP3vMapCDwfEC7fRxkBpzD5ODnhXEwH79U8a7RU1Ic/export?format=xlsx");

  const a = await f.arrayBuffer();
  const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
  if (wb) {
    hiddenLoader();
  }
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

const topLabels = {
  id: "topLabels",
  afterDatasetsDraw(chart, args, pluginOpt) {
    const { ctx, scales: { x, y } } = chart;
    chart.data.datasets[0].data.forEach((datapoint, index) => {
      const datasetArray = [];
      chart.data.datasets.forEach((dataset) => {
        if(dataset.data[index] != undefined && !isNaN(dataset.data[index])){
          datasetArray.push(dataset.data[index]);
        }
      });

      function totalSum(total, values) {
        return total + values;
      }

      let sum = datasetArray.reduce(totalSum, 0);
      ctx.font = "bold";
      ctx.fillStyle = "#808080";
      ctx.textAlign = "center";

      ctx.fillText(sum.toFixed(1), x.getPixelForValue(index), 75);
    });
  }
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

        const totalArray = [buildValue($("#total1")), buildValue($("#total2")), buildValue($("#total3")) ]

        totalArray.forEach((data, index) => {
          ctx.font = "bold";
          ctx.fillStyle = "#808080";
          ctx.textAlign = "center";
          ctx.fillText(data, x.getPixelForValue(index), y.chart.chartArea.bottom - y.chart.chartArea.height + 6);
        })
      }}]
  };
};


HydrogenCalc.fn.getVal = function(sheet, addr) {
  return this.sheets[sheet ?? "SMR"].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
};

HydrogenCalc.fn.drawChart = function() {
  self = this;
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
    }
  ];

  this.chart = new Chart($("#chart_container")[0], chartOpts);
};

$("#taxCredit, #carbonPrice, #carbon, #electricity, #gas").keypress(function(e) {
  if (e.which == 13) {
    $("#taxCredit, #carbonPrice, #carbon, #electricity, #gas").blur();
  }
});

function buildValue(element) {
  const value = element.val();
  return value.replace(/^[-+]?[0-9]*[.,]?[0-9]+$/g, "").replace("$", "");
}

$("#gas, #electricity, #carbon, #carbonPrice, #taxCredit").change(function() {
  setTimeout(function() {
    chartOpts.data.datasets = [{
      label: "Power Export",
      data: [
        buildValue($("#pe1")),
        buildValue($("#pe2")),
        buildValue($("#pe3"))
      ],
      backgroundColor: "rgba(75,97,182,0.2)",
      borderColor: [
        "rgb(72,15,217)"
      ],
      borderWidth: 1
    }, {
      label: "CAPEX",
      data: [
        buildValue($("#cap1")),
        buildValue($("#cap2")),
        buildValue($("#cap3"))
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
        buildValue($("#fo1")),
        buildValue($("#fo2")),
        buildValue($("#fo3"))
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
        buildValue($("#natGas1")),
        buildValue($("#natGas2")),
        buildValue($("#natGas3"))
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
      data: [buildValue($("#el1")),
        buildValue($("#el2")),
        buildValue($("#el3"))
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
        buildValue($("#wat1")),
        buildValue($("#wat2")),
        buildValue($("#wat3"))
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
        buildValue($("#co1")),
        buildValue($("#co2")),
        buildValue($("#co3"))
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
        buildValue($("#ct1")),
        buildValue($("#ct2")),
        buildValue($("#ct3"))      ],
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
        buildValue($("#ccr1")),
        buildValue($("#ccr2")),
        buildValue($("#ccr3"))
      ],
      backgroundColor: [
        "rgba(25, 9, 232, 0.2)"
      ],
      borderColor: [
        "rgba(25, 9, 232, 1)"
      ],
      borderWidth: 1
    }
    ];
    $("#canvas_wr").html(""); //remove canvas from container
    $("#canvas_wr").html("   <canvas id=\"chart_container\" height=\"200px\"></canvas>"); //add it back to the container
    this.chart = new Chart($("#chart_container")[0], chartOpts);
  }, 3000);
});


//AMMONIA CALC ________________________________________________________________________________________________________

AmmoniaCalc.fn.init = async function() {
  self2 = this;
  const f = await fetch("https://docs.google.com/spreadsheets/d/136gX-lPlD2fMbYbxCJmtu_f762j5SfbAuUohJT_MdkU/export?format=xlsx");
  const a = await f.arrayBuffer();
  const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
  $("#body").css("overflow", "hidden");
  if (wb) {
    hiddenLoader();
  }
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
        const totalArray = [buildValue($("#amtotal1")), buildValue($("#amtotal2"))]
        totalArray.forEach((data, index) => {
          ctx.font = "bold";
          ctx.fillStyle = "#808080";
          ctx.textAlign = "center";

          ctx.fillText(data, x.getPixelForValue(index), y.chart.chartArea.bottom - y.chart.chartArea.height + 6);
        })
      }}]
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
    }
  ];

  this.chart2 = new Chart($("#chart_container2")[0], chartOpts2);
};

$("#taxCreditAm, #carbonPriceAm, #carbonAm, #electricityAm, #gasAm, #elExportAm").keypress(function(e) {
  if (e.which == 13) {
    $("#taxCreditAm, #carbonPriceAm, #carbonAm, #electricityAm, #gasAm, #elExportAm").blur();
  }
});

$("#gasAm, #electricityAm, #carbonAm, #carbonPriceAm, #taxCreditAm, #elExportAm").change(function() {
  function buildValue(element) {
    const value = element.val();
    return value.replace(/^[-+]?[0-9]*[.,]?[0-9]+$/g, "").replace("$", "");
  }

  setTimeout(function() {
    chartOpts2.data.datasets = [{
      label: "CAPEX",
      data: [
        buildValue($("#amcap1")),
        buildValue($("#amcap2"))
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
        buildValue($("#amfo1")),
        buildValue($("#amfo2"))
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
        buildValue($("#amng1")),
        buildValue($("#amng2"))
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
        buildValue($("#amel1")), buildValue($("#amel2"))
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
        buildValue($("#amwat1")),
        buildValue($("#amwat2"))
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
        buildValue($("#amco1")),
        buildValue($("#amco2"))
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
        buildValue($("#amct1")),
        buildValue($("#amct2"))
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
        buildValue($("#amccr1")),
        buildValue($("#amccr2"))
      ],
      backgroundColor: [
        "rgba(25, 9, 232, 0.2)"
      ],
      borderColor: [
        "rgba(25, 9, 232, 1)"
      ],
      borderWidth: 1
    }];
    $("#canvas_wr2").html(""); //remove canvas from container
    $("#canvas_wr2").html("   <canvas id=\"chart_container2\" height=\"200px\"></canvas>"); //add it back to the container
    this.chart2 = new Chart($("#chart_container2")[0], chartOpts2);
  }, 3000);
});


//GREEN CALC ________________________________________________________________________________________________________

function hiddenLoader() {
  setTimeout(function() {
    $("#loaderMain").css("display", "none");
    $("#body").css("overflow", "auto");
  }, 2000);
}

GreenHydrogenCalc.fn.init = async function() {
  self3 = this;
  const f = await fetch("https://docs.google.com/spreadsheets/d/1QqSpCssFTjQUKgWyHe4BTNVhN7tIFs8ufOBlbS5Lsfw/export?format=xlsx");
  const a = await f.arrayBuffer();
  const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
  $("#body").css("overflow", "hidden");
  if (wb) {
    hiddenLoader();
  }
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
      labels: ["PEM", "AEL"],
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
        chart.data.datasets[0].data.forEach((datapoint, index) => {
          const datasetArray = [];
          chart.data.datasets.forEach((dataset) => {
            if(dataset.data[index] != undefined && !isNaN(dataset.data[index])){
              datasetArray.push(dataset.data[index]);
            }
          });

          function totalSum(total, values) {
            return total + values;
          }

          let sum = datasetArray.reduce(totalSum, 0);
          ctx.font = "bold";
          ctx.fillStyle = "#808080";
          ctx.textAlign = "center";
          ctx.fillText(sum.toFixed(1), x.getPixelForValue(index), y.chart.chartArea.bottom - y.chart.chartArea.height + 6);
        });
      }
    }]
  };
};


GreenHydrogenCalc.fn.getVal = function(sheet, addr) {
  return this.sheets[sheet ?? "NH3PEM"].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
};


GreenHydrogenCalc.fn.drawChart = function() {
  self3 = this;
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
      self3.getVal("H2AEL", "C89")
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
        self3.getVal("H2AEL", "C88")
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
        self3.getVal("H2AEL", "C87")
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
        self3.getVal("H2AEL", "C86")
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
        self3.getVal("H2AEL", "C85")
      ],
      backgroundColor: [
        "rgba(68,134,60,0.2)"
      ],
      borderColor: [
        "rgb(14,227,7)"
      ],
      borderWidth: 1
    }
  ];

  this.chart3 = new Chart($("#chart_container3")[0], chartOpts3);

  chartOpts4.data.datasets = [{
    label: "CAPEX",
    data: [
      self3.getVal("NH3PEM", "G89"),
      self3.getVal("NH3AEL", "C89")
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
        self3.getVal("NH3AEL", "C88")
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
        self3.getVal("NH3AEL", "C87")
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
        self3.getVal("NH3AEL", "C86")
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
        self3.getVal("NH3AEL", "C85")
      ],
      backgroundColor: [
        "rgba(68,134,60,0.2)"
      ],
      borderColor: [
        "rgb(14,227,7)"
      ],
      borderWidth: 1
    }
  ];

  this.chart4 = new Chart($("#chart_container4")[0], chartOpts4);
};

$("#gasGr, #electricityGr, #capexGr, #opexGr, #electrolyzerEfGr, #ptcTaxCredit").keypress(function(e) {
  if (e.which == 13) {
    $("#gasGr, #electricityGr, #capexGr, #opexGr, #electrolyzerEfGr, #ptcTaxCredit").blur();
  }
});

$("#gasGr, #electricityGr, #capexGr, #opexGr, #electrolyzerEfGr, #ptcTaxCredit").change(function() {
  function buildValue(element) {
    const value = element.val();
    return value.replace(/^[-+]?[0-9]*[.,]?[0-9]+$/g, "").replace("$", "");
  }

  setTimeout(function() {
    chartOpts3.data.datasets = [{
      label: "CAPEX",
      data: [
        parseFloat(buildValue($("#gr1cap1"))),
        parseFloat(buildValue($("#gr1cap2")))
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
        parseFloat(buildValue($("#gr1fo1"))),
        parseFloat(buildValue($("#gr1fo2")))
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
        parseFloat(buildValue($("#gr1el1"))),
        parseFloat(buildValue($("#gr1el2")))
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
        parseFloat(buildValue($("#gr1wat1"))),
        parseFloat(buildValue($("#gr1wat2")))
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
        parseFloat(buildValue($("#gr1TaxCr1"))),
        parseFloat(buildValue($("#gr1TaxCr2")))
      ],
      backgroundColor: [
        "rgba(25, 9, 232, 0.2)"
      ],
      borderColor: [
        "rgba(25, 9, 232, 1)"
      ],
      borderWidth: 1
    }
    ];

    chartOpts4.data.datasets = [{
      label: "CAPEX",
      data: [
        parseFloat(buildValue($("#gr2cap1"))),
        parseFloat(buildValue($("#gr2cap2")))
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
        parseFloat(buildValue($("#gr2fo1"))),
        parseFloat(buildValue($("#gr2fo2")))
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
        parseFloat(buildValue($("#gr2el1"))),
        parseFloat(buildValue($("#gr2el2")))
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
        parseFloat(buildValue($("#gr2wat1"))),
        parseFloat(buildValue($("#gr2wat2")))
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
        parseFloat(buildValue($("#gr2TaxCr1"))),
        parseFloat(buildValue($("#gr2TaxCr2")))
      ],
      backgroundColor: [
        "rgba(25, 9, 232, 0.2)"
      ],
      borderColor: [
        "rgba(25, 9, 232, 1)"
      ],
      borderWidth: 1
    }
    ];


    $("#canvas_wr3").html(""); //remove canvas from container
    $("#canvas_wr3").html("   <canvas id=\"chart_container3\" height=\"200px\"></canvas>"); //add it back to the container
    this.chart3 = new Chart($("#chart_container3")[0], chartOpts3);

    $("#canvas_wr4").html(""); //remove canvas from container
    $("#canvas_wr4").html("   <canvas id=\"chart_container4\" height=\"200px\"></canvas>"); //add it back to the container
    this.chart4 = new Chart($("#chart_container4")[0], chartOpts4);
  }, 3000);
});
