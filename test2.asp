
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<meta http-equiv="X-UA-Compatible" content="ie=edge">
		<!-- <meta content="60; url=http://ufa-qos01ow/vsp/main2.asp" http-equiv=refresh> -->
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<script type="text/javascript" src="js/highstock.js"></script>
		<!-- <script type="text/javascript" src="js/highcharts.js"></script>-->
		<script type="text/javascript" src="js/themes/gray.js"></script>
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->
		<script type="text/javascript">
		
			var chart5;
			var chartE;
			var chart6;
			var chartF;
			var chart7;
			var chartG;

			// Первый график
			$(document).ready(function() {
				
				chart7 = new Highcharts.Chart({
					chart: {
						renderTo: 'container7',
						// defaultSeriesType: 'column'
						type: 'line'
					},
					colors: ['#66FFFF', '#FFFF66'], //'#FF66FF'
					credits: {enabled: false},
				/*	legend:  {enabled: false},
					tooltip: {enabled: false},*/
					title:   {align: 'right', text: 'Клиентов в очереди'},
					xAxis: [{
						max: Date.UTC(2018, 07, 30, 13, 39),
						type: 'datetime',
						dateTimeLabelFormats: { // don't display the dummy year
							hour: '%H:%M'
						}
					}],
					yAxis: [
					{
					    min: 0,
						title: {
							text: null
						},
						lineColor: '#66FFFF',
						allowDecimals: false,
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					}
					],
					plotOptions: {
						line: {
							dataLabels: {
								enabled: true,
								formatter: function() {
									return this.y > 0  ? this.y : null; 
								}
							},
							enableMouseTracking: false
						},
                        series: {
                            enableMouseTracking: false,
                            marker: {
                                enabled: true
                            }
                        }
					},
					legend: {
						enabled : false,
						layout: 'horizontal',
						floating: true,
						backgroundColor: '#363636',
						align: 'left',
						verticalAlign: 'top',
						x: 4,
						y: -8,
						borderWidth: 0
					},
					series: [
				{
						name: 'ДПП'
, type: 'scatter', 
data: [  {color: null, marker: { enabled: false },
    //fillColor: '#FF0000', lineColor: '#FF0000', radius: 2},  
x: Date.UTC(2018, 7, 30, 12, 19), y: 1 } ] 
					}, {
						name: 'ТСП'
, data: [[Date.UTC(2018, 7, 30, 12, 23), 0],
[Date.UTC(2018, 7, 30, 12, 35), 0],
[Date.UTC(2018, 7, 30, 12, 46), 0],
[Date.UTC(2018, 7, 30, 12, 57), 0],
[Date.UTC(2018, 7, 30, 13, 8), 0]]

					}
					


 /*                       {
						name: 'ДПП'
, type: 'scatter', data: [  {color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2},  x: Date.UTC(2018, 7, 30, 11, 39), y: -1 } ] 
					}, 
                    {
						name: 'ТСП'
, data: [[Date.UTC(2018, 7, 30, 12, 0), 0],
[Date.UTC(2018, 7, 30, 12, 3), 1],
[Date.UTC(2018, 7, 30, 12, 5), 3],
[Date.UTC(2018, 7, 30, 12, 6), 4],
[Date.UTC(2018, 7, 30, 12, 7), 2],
[Date.UTC(2018, 7, 30, 12, 8), 0],
[Date.UTC(2018, 7, 30, 12, 9), 1],
[Date.UTC(2018, 7, 30, 12, 10), 0],
[Date.UTC(2018, 7, 30, 12, 11), 1],
[Date.UTC(2018, 7, 30, 12, 12), 0],
[Date.UTC(2018, 7, 30, 12, 23), 0],
[Date.UTC(2018, 7, 30, 12, 35), 0],
[Date.UTC(2018, 7, 30, 12, 46), 0],
[Date.UTC(2018, 7, 30, 12, 57), 0],
[Date.UTC(2018, 7, 30, 13, 8), 0]]

					}*/
					]
				});


				chartG = new Highcharts.Chart({
					chart:   {renderTo: 'containerG', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartG.renderer.circle(150, 150, 90).attr({
					fill: '#00FF00',
					stroke: '#00FF00'
				}).add();
				chartG.renderer.image('q.gif', 75, 75, 150, 150).add();
				
			});

		</script>
		
	<style type="text/css">
	<!--
	BODY {
		margin: 0px;
		background-color: #242424;
	}
TABLE {
	margin-bottom: 0px;
	margin-top: 0px;
}
TD {
	padding-top: 1px;
	padding-bottom: 1px;
	text-align: center;
	color: #FFFFFF;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
}
TD.A {
	border: solid 1px #4572A7
}
TD.Head {
	color: #000000;
	font-size: 28pt;
}
TD.Txt {
	color: #FFFFFF;
	font-size: 48pt;
	font-weight: 700;
}
	
	-->
	</style>
</head>
<body>
<div align="center" valign="top">

    <div id="container7"  style="width: 940px; height: 320px; margin: 0 auto"></div>

</div>

</body>
</html>

<!-- Разработка: Машков А.В. -->
<!-- Для вывода графиков используется библиотека Highcharts JS - http://highsoft.com/ -->
