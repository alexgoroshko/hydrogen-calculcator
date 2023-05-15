function HydrogenCalc() {
    this.$ = jQuery;
    this.chart;
    this.xlsx = XLSX;
    this.sheets = {};
    this.tabs = ['#Dashboard', '#Assumptions', '#BaseCase', '#SMR', '#ATR'];
    this.init();
}

HydrogenCalc.fn = HydrogenCalc.prototype;

let self = null;
let chartOpts;

function formulaParse(formula, sheetNames) {
    let cleanFormula = formula.split('$').join('');
    //Removing spaces in sheet names
    sheetNames.forEach(sheetName => {
        cleanFormula = cleanFormula.split(sheetName).join(`#${sheetName.split(' ').join('')}`);
    });
    cleanFormula = cleanFormula.split('\'').join('');
    cleanFormula = cleanFormula.split('>').join(' > ');
    cleanFormula = cleanFormula.split('<').join(' < ');
    return cleanFormula;
}

function splitName(name) {
    return name.split(' ').join('');
}

HydrogenCalc.fn.init = async function() {
    self = this;
    const f = await fetch('https://docs.google.com/spreadsheets/d/1lcpnDp8JhwKlMBLvV73h9l2hxyLV2pcoRBKCcDQe4AU/export?format=xlsx');
    const a = await f.arrayBuffer();
    const wb = this.xlsx.read(a, { cellFormula: true, cellNF: true });
    const hydrogenData = {};
    Object.keys(wb.Sheets).forEach(name => {
        Object.keys(wb.Sheets[name]).forEach(cell => {
            if (!hydrogenData[splitName(name)]) {
                hydrogenData[splitName(name)] = {};
            }
            // We only need cells
            if (cell[0] === '!') {
                return;
            }
            if (wb.Sheets[name][cell].f) {
                hydrogenData[splitName(name)][cell] = {
                    format: wb.Sheets[name][cell].z || '',
                    formula: formulaParse(wb.Sheets[name][cell].f, wb.SheetNames),
                    value: wb.Sheets[name][cell].v
                };
            } else {
                hydrogenData[splitName(name)][cell] = {
                    format: wb.Sheets[name][cell].z || '',
                    value: wb.Sheets[name][cell].v
                };
            }
            if (hydrogenData[splitName(name)][cell].format === 'General') {
                hydrogenData[splitName(name)][cell].format = '';
            }
        });
    });

    console.log(hydrogenData)

    $(self.tabs.join(',')).calx({
        data: hydrogenData,
        onAfterCalculate: function() {
            if (self.chart) {
                self.chart.data.datasets[0].data = [];

                self.chart.update();
            }
        }
    });

    self.tabs.map(function(tab) {
        self.sheets[tab.replace('#', '')] = self.$(tab).calx('getSheet');
    });

    self.drawChart();
};

HydrogenCalc.fn.getDefaultChartOpts = function() {
    return {
        type: 'bar',
        data: {
            labels: ['Reference Case', 'SMR +90% CCS', 'ATR + GHR'],
            datasets: []
        },
        options: {
            scales: {
                x: {
                    stacked: true
                },
                y: {
                    stacked: true
                }
            }
        }
    };
};


HydrogenCalc.fn.getVal = function(sheet, addr) {
    return this.sheets[sheet ?? 'SMR'].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
};


HydrogenCalc.fn.drawChart = function() {
    self = this;
    chartOpts = self.getDefaultChartOpts();
    chartOpts.data.datasets = [{
        label: 'Power Export',
        data: [
            self.getVal('BaseCase', 'H117'),
            self.getVal('BaseCase', 'I117'),
            self.getVal('BaseCase', 'J117')
        ],
        backgroundColor: 'rgba(255, 99, 132, 0.2)',
        borderColor: [
            'rgba(255, 99, 132, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'CAPEX',
        data: [
            self.getVal('BaseCase', 'H118'),
            self.getVal('BaseCase', 'I118'),
            self.getVal('BaseCase', 'J118')
        ],
        backgroundColor: [
            'rgba(54, 162, 235, 0.2)'
        ],
        borderColor: [
            'rgba(54, 162, 235, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'Fixed OPEX',
        data: [
            self.getVal('BaseCase', 'H119'),
            self.getVal('BaseCase', 'I119'),
            self.getVal('BaseCase', 'J119')
        ],
        backgroundColor: [
            'rgba(255, 206, 86, 0.2)'
        ],
        borderColor: [
            'rgba(255, 206, 86, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'Feedstock',
        data: [
            self.getVal('BaseCase', 'H120'),
            self.getVal('BaseCase', 'I120'),
            self.getVal('BaseCase', 'J120')
        ],
        backgroundColor: [
            'rgba(75, 192, 192, 0.2)'
        ],
        borderColor: [
            'rgba(75, 192, 192, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'Fuel',
        data: [
            self.getVal('BaseCase', 'H121'),
            self.getVal('BaseCase', 'I121'),
            self.getVal('BaseCase', 'J121')
        ],
        backgroundColor: [
            'rgba(153, 102, 255, 0.2)'
        ],
        borderColor: [
            'rgba(153, 102, 255, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'Electricity',
        data: [
            self.getVal('BaseCase', 'H122'),
            self.getVal('BaseCase', 'I122'),
            self.getVal('BaseCase', 'J122')
        ],
        backgroundColor: [
            'rgba(255, 159, 64, 0.2)'
        ],
        borderColor: [
            'rgba(255, 159, 64, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'Water',
        data: [
            self.getVal('BaseCase', 'H123'),
            self.getVal('BaseCase', 'I123'),
            self.getVal('BaseCase', 'J123')
        ],
        backgroundColor: [
            'rgba(255, 99, 132, 0.2)'
        ],
        borderColor: [
            'rgba(255, 99, 132, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'CO2 T&S',
        data: [
            self.getVal('BaseCase', 'H124'),
            self.getVal('BaseCase', 'I124'),
            self.getVal('BaseCase', 'J124')
        ],
        backgroundColor: [
            'rgba(80, 99, 132, 0.2)'
        ],
        borderColor: [
            'rgba(55, 99, 132, 1)'
        ],
        borderWidth: 1
    }, {
        label: 'Hydrogen Distribution',
        data: [
            self.getVal('BaseCase', 'H125'),
            self.getVal('BaseCase', 'I125'),
            self.getVal('BaseCase', 'J125')
        ],
        backgroundColor: 'rgb(120, 120, 120)',
        borderWidth: 1
    }, {
        label: 'Carbon Price',
        data: [
            self.getVal('BaseCase', 'H126'),
            self.getVal('BaseCase', 'I126'),
            self.getVal('BaseCase', 'J126')
        ],
        backgroundColor: [
            'rgba(25, 9, 232, 0.2)'
        ],
        borderColor: [
            'rgba(25, 9, 232, 1)'
        ],
        borderWidth: 1
    }];

    this.chart = new Chart($('#chart_container')[0], chartOpts);
};

$('#taxCredit, #carbonPrice, #carbon, #electricity, #gas').keypress(function(e) {
    if (e.which == 13) {
        $('#taxCredit, #carbonPrice, #carbon, #electricity, #gas').blur();
    }
});

$('#gas, #electricity, #carbon, #carbonPrice, #taxCredit').change(function() {
    function buildValue(element) {
        console.log(element)
        const value = element.val();
        console.log(value)
        const res = value.replace(/[^.\d]/g, '');
        return value.includes('(') ? `-${res}` : res;
    }

    setTimeout(function() {
        chartOpts.data.datasets = [{
            label: 'Power Export',
            data: [
                buildValue($('#pe1')),
                buildValue($('#pe2')),
                buildValue($('#pe3'))
            ],
            backgroundColor: 'rgba(255, 99, 132, 0.2)',
            borderColor: [
                'rgba(255, 99, 132, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'CAPEX',
            data: [
                buildValue($('#cap1')),
                buildValue($('#cap2')),
                buildValue($('#cap3'))
            ],
            backgroundColor: [
                'rgba(54, 162, 235, 0.2)'
            ],
            borderColor: [
                'rgba(54, 162, 235, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'Fixed OPEX',
            data: [
                buildValue($('#fo1')),
                buildValue($('#fo2')),
                buildValue($('#fo3'))
            ],
            backgroundColor: [
                'rgba(255, 206, 86, 0.2)'
            ],
            borderColor: [
                'rgba(255, 206, 86, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'Feedstock',
            data: [
                buildValue($('#fs1')),
                buildValue($('#fs2')),
                buildValue($('#fs3'))
            ],
            backgroundColor: [
                'rgba(75, 192, 192, 0.2)'
            ],
            borderColor: [
                'rgba(75, 192, 192, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'Fuel',
            data: [
                buildValue($('#fuel1')),
                buildValue($('#fuel2')),
                buildValue($('#fuel3'))
            ],
            backgroundColor: [
                'rgba(153, 102, 255, 0.2)'
            ],
            borderColor: [
                'rgba(153, 102, 255, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'Electricity',
            data: [
                buildValue($('#el1')),
                buildValue($('#el2')),
                buildValue($('#el3'))
            ],
            backgroundColor: [
                'rgba(255, 159, 64, 0.2)'
            ],
            borderColor: [
                'rgba(255, 159, 64, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'Water',
            data: [
                buildValue($('#wat1')),
                buildValue($('#wat2')),
                buildValue($('#wat3'))
            ],
            backgroundColor: [
                'rgba(31, 2, 217, 0.2);'
            ],
            borderColor: [
                'rgba(31, 2, 217, 1);'
            ],
            borderWidth: 1
        }, {
            label: 'CO2 T&S',
            data: [
                buildValue($('#co1')),
                buildValue($('#co2')),
                buildValue($('#co3'))
            ],
            backgroundColor: [
                'rgba(80, 99, 132, 0.2)'
            ],
            borderColor: [
                'rgba(55, 99, 132, 1)'
            ],
            borderWidth: 1
        }, {
            label: 'Hydrogen Distribution',
            data: [
                buildValue($('#hd1')),
                buildValue($('#hd2')),
                buildValue($('#hd3'))
            ],
            backgroundColor: 'rgb(120, 120, 120)',
            borderWidth: 1
        }, {
            label: 'Carbon Price',
            data: [
                buildValue($('#cp1')),
                buildValue($('#cp2')),
                buildValue($('#cp3'))
            ],
            backgroundColor: [
                'rgba(25, 9, 232, 0.2)'
            ],
            borderColor: [
                'rgba(25, 9, 232, 1)'
            ],
            borderWidth: 1
        }];
        $('#canvas_wr').html(''); //remove canvas from container
        $('#canvas_wr').html('   <canvas id="chart_container" height="200px"></canvas>'); //add it back to the container
        this.chart = new Chart($('#chart_container')[0], chartOpts);
    }, 2000);
});
