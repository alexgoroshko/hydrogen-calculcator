function HydrogenCalc() {
    this.$ = jQuery;
    this.chart;
    this.sheets = {};
    this.tabs = ['#Dashboard', '#Assumptions', '#BaseCase', '#SMR', '#ATR']
    this.init();
}

HydrogenCalc.fn = HydrogenCalc.prototype;

let self = null
let chartOpts

HydrogenCalc.fn.init = function () {
    self = this;

    $(self.tabs.join(',')).calx({
        data: hydrogenData,
        onAfterCalculate: function () {
            if (self.chart) {
                self.chart.data.datasets[0].data = [
                ];

                self.chart.update();
            }
        }
    });

    self.tabs.map(function (tab) {
        self.sheets[tab.replace('#', '')] = self.$(tab).calx('getSheet');
    });

    self.drawChart();
}

HydrogenCalc.fn.getDefaultChartOpts = function () {
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
            },
        }
    }
}



HydrogenCalc.fn.getVal = function (sheet, addr) {
    return this.sheets[sheet ?? 'SMR'].cells[addr] ? this.sheets[sheet].cells[addr].getValue() : 0;
}


HydrogenCalc.fn.drawChart = function () {
    self = this;
    chartOpts = self.getDefaultChartOpts();
    chartOpts.data.datasets = [{
        label: 'Power Export',
        data: [
            self.getVal('Dashboard', 'A1'),
            self.getVal('Dashboard', 'A2'),
            self.getVal('Dashboard', 'B6'),
        ],
        backgroundColor: 'rgba(255, 99, 132, 0.2)',
        borderColor: [
            'rgba(255, 99, 132, 1)',
        ],
        borderWidth: 1,
    }, {
        label: 'CAPEX',
        data: [
            self.getVal('Dashboard', 'F6'),
            self.getVal('Dashboard', 'J6'),
            self.getVal('Dashboard', 'N6'),
        ],
        backgroundColor: [
            'rgba(54, 162, 235, 0.2)'
        ],
        borderColor: [
            'rgba(54, 162, 235, 1)',
        ],
        borderWidth: 1
    }, {
        label: 'Fixed OPEX',
        data: [
            self.getVal('Dashboard', 'R6'),
            self.getVal('Dashboard', 'B8'),
            self.getVal('Dashboard', 'F8'),
        ],
        backgroundColor: [
            'rgba(255, 206, 86, 0.2)',
        ],
        borderColor: [
            'rgba(255, 206, 86, 1)',
        ],
        borderWidth: 1
    }, {
        label: 'Feedstock',
        data: [
            self.getVal('Dashboard', 'L41'),
            self.getVal('Dashboard', 'B9'),
            self.getVal('Dashboard', 'F9'),
        ],
        backgroundColor: [
            'rgba(75, 192, 192, 0.2)',
        ],
        borderColor: [
            'rgba(75, 192, 192, 1)',
        ],
        borderWidth: 1
    }, {
        label: 'Fuel',
        data: [
            self.getVal('Dashboard', 'B41'),
            self.getVal('Dashboard', 'B10'),
            self.getVal('Dashboard', 'F10'),
        ],
        backgroundColor: [
            'rgba(153, 102, 255, 0.2)',
        ],
        borderColor: [
            'rgba(153, 102, 255, 1)',
        ],
        borderWidth: 1
    }, {
        label: 'Electricity',
        data: [
            self.getVal('Dashboard', 'L19'),
            self.getVal('Dashboard', 'J10'),
            self.getVal('Dashboard', 'N10'),
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
            self.getVal('Dashboard', 'R10'),
            self.getVal('Dashboard', 'B11'),
            self.getVal('Dashboard', 'F11'),
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
            self.getVal('Dashboard', 'B12'),
            self.getVal('Dashboard', 'F12'),
            self.getVal('Dashboard', 'B13'),
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
            self.getVal('Dashboard', 'F13'),
            self.getVal('Dashboard', 'J13'),
            self.getVal('Dashboard', 'N13'),
        ],
        backgroundColor: 'rgb(120, 120, 120)',
        borderWidth: 1
    }, {
        label: 'Carbon Price',
        data: [
            self.getVal('Dashboard', 'R13'),
            self.getVal('Dashboard', 'B14'),
            self.getVal('Dashboard', 'F14'),
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
}

$("#taxCredit, #carbonPrice, #carbon, #electricity, #gas").keypress(function (e) {
    if(e.which == 13) {
        $("#taxCredit, #carbonPrice, #carbon, #electricity, #gas").blur();
    }
})

$("#gas, #electricity, #carbon, #carbonPrice, #taxCredit").change(function () {
    function buildValue(element) {
        const value = element.val();
        const res = value.replace(/[^.\d]/g, '');
        return value.includes('(') ? `-${res}` : res
    }

    setTimeout(function () {
        chartOpts.data.datasets = [{
            label: 'Power Export',
            data: [
                buildValue($('#pe1')),
                buildValue($('#pe2')),
                buildValue($('#pe3')),
            ],
            backgroundColor: 'rgba(255, 99, 132, 0.2)',
            borderColor: [
                'rgba(255, 99, 132, 1)',
            ],
            borderWidth: 1,
        }, {
            label: 'CAPEX',
            data: [
                buildValue($('#cap1')),
                buildValue($('#cap2')),
                buildValue($('#cap3')),
            ],
            backgroundColor: [
                'rgba(54, 162, 235, 0.2)'
            ],
            borderColor: [
                'rgba(54, 162, 235, 1)',
            ],
            borderWidth: 1
        }, {
            label: 'Fixed OPEX',
            data: [
                buildValue($('#fo1')),
                buildValue($('#fo2')),
                buildValue($('#fo3')),
            ],
            backgroundColor: [
                'rgba(255, 206, 86, 0.2)',
            ],
            borderColor: [
                'rgba(255, 206, 86, 1)',
            ],
            borderWidth: 1
        }, {
            label: 'Feedstock',
            data: [
                buildValue($('#fs1')),
                buildValue($('#fs2')),
                buildValue($('#fs3')),
            ],
            backgroundColor: [
                'rgba(75, 192, 192, 0.2)',
            ],
            borderColor: [
                'rgba(75, 192, 192, 1)',
            ],
            borderWidth: 1
        }, {
            label: 'Fuel',
            data: [
                buildValue($('#fuel1')),
                buildValue($('#fuel2')),
                buildValue($('#fuel3')),
            ],
            backgroundColor: [
                'rgba(153, 102, 255, 0.2)',
            ],
            borderColor: [
                'rgba(153, 102, 255, 1)',
            ],
            borderWidth: 1
        }, {
            label: 'Electricity',
            data: [
                buildValue($('#el1')),
                buildValue($('#el2')),
                buildValue($('#el3')),
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
                buildValue($('#wat3')),
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
                buildValue($('#co3')),
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
                buildValue($('#hd3')),
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