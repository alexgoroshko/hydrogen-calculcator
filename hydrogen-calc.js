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
    setTimeout(function () {
        chartOpts.data.datasets = [{
            label: 'Power Export',
            data: [
                $('#pe1').val().replace(/[^.\d]/g, ''),
                $('#pe2').val().replace(/[^.\d]/g, ''),
                $('#pe3').val().replace(/[^.\d]/g, ''),
            ],
            backgroundColor: 'rgba(255, 99, 132, 0.2)',
            borderColor: [
                'rgba(255, 99, 132, 1)',
            ],
            borderWidth: 1,
        }, {
            label: 'CAPEX',
            data: [
                $('#cap1').val().replace(/[^.\d]/g, ''),
                $('#cap2').val().replace(/[^.\d]/g, ''),
                $('#cap3').val().replace(/[^.\d]/g, ''),
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
                $('#fo1').val().replace(/[^.\d]/g, ''),
                $('#fo2').val().replace(/[^.\d]/g, ''),
                $('#fo3').val().replace(/[^.\d]/g, ''),
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
                $('#fs1').val().replace(/[^.\d]/g, ''),
                $('#fs2').val().replace(/[^.\d]/g, ''),
                $('#fs3').val().replace(/[^.\d]/g, ''),
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
                $('#fuel1').val().replace(/[^.\d]/g, ''),
                $('#fuel2').val().replace(/[^.\d]/g, ''),
                $('#fuel3').val().replace(/[^.\d]/g, ''),
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
                $('#el1').val().replace(/[^.\d]/g, ''),
                $('#el2').val().replace(/[^.\d]/g, ''),
                $('#el3').val().replace(/[^.\d]/g, ''),
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
                $('#wat1').val().replace(/[^.\d]/g, ''),
                $('#wat2').val().replace(/[^.\d]/g, ''),
                $('#wat3').val().replace(/[^.\d]/g, ''),
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
                $('#co1').val().replace(/[^.\d]/g, ''),
                $('#co2').val().replace(/[^.\d]/g, ''),
                $('#co3').val().replace(/[^.\d]/g, ''),
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
                $('#hd1').val().replace(/[^.\d]/g, ''),
                $('#hd2').val().replace(/[^.\d]/g, ''),
                $('#hd3').val().replace(/[^.\d]/g, ''),
            ],
            backgroundColor: 'rgb(120, 120, 120)',
            borderWidth: 1
        }, {
            label: 'Carbon Price',
            data: [
                $('#cp1').val().replace(/[^.\d]/g, ''),
                $('#cp2').val().replace(/[^.\d]/g, ''),
                $('#cp3').val().replace(/[^.\d]/g, '')
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