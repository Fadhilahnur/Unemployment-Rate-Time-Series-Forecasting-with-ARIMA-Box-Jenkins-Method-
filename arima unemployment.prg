'==============
'Title: This is a program to forecast unemployment in Malaysia using Box-Jenkins Method
'Group 2 : Unemployment Rate in Malaysia
'Author: Fadhilah Nur Binti ismail
'Matric ID: A167808
'Date: 9 June 2020 (Tuesday)
'==============

close @all 'untuk clear everything files before
logmode l

%path = @runpath
cd %path
%xlsdata = "arima.xlsx"
%pathxlsdata = %path+%xlsdata
%file = "arima"

'Declare tahun
%dat_start = "2010M01"
%dat_end = "2019M12"

%fcst_start = "2020M01"
%fcst_end = "2020M09"

%smpl_start ="2010M01"
%smpl_end = "2019M12"

'Create workfile
'wfcreate(wf={%file}) m %dat_start %fcst_end
wfcreate m %dat_start %fcst_end	
pagerename Untitled {%file}

string sheet1 = "GROUP_2"

for %a {sheet1}
import %pathxlsdata range=%a colhead=1 na="#N/A" @freq m @id @date(series01) @destid @date @smpl @all
next

logmsg Fadhilah has done importing {%file}

''==============
'1) Identification
''==============
smpl %smpl_start %smpl_end

'Data transformation
series unrate =(un/lf)*100

%a = "UNrate"    'pilih variable yg nak guna
graph graphl{%a}.line log({%a}) 
freeze(correll{%a}) log({%a}).correl    'untuk tengok graph correlogram

series l{%a} = log({%a})
series dlog{%a} = l{%a}-l{%a}(-1) 'jadikan series stationary
graph graphdl{%a}.line dlog({%a})
freeze(correll{%a}d1) log({%a}).correl(d=1) 'd=1 maksudnya first difference

logmsg Fadhilah has completed analysis on identification 
 
''==============
'2) Estimation
''==============
'Set berapa lag yang perlu untuk remove MA term dengan AR term
'Letak lag sebagai object variable
'a sebagai AR term
'b sebagai MA term
'pilih AR dan MA yang lepas boundaries sahaja

!a = 1
!b = 1
equation arima{%a}!a!b.ls(optmethod=opg) dlog({%a}) c ar(!a) ma(!b)

!a = 2
!b = 1
equation arima{%a}!a!b.ls(optmethod=opg) dlog({%a}) c ar(!a) ma(!b)

logmsg Fadhilah has completed analysis on estimation

'To choose the best estimation, compare by 
'1. The most significant coefficients
'2. The least volatility (shown by sigmasq)
'3. Lowest AIC (Akaike) and SBIC (schwartz)
'4. Highest Adjusted R-squared

''==============
'3) Diagnostic
''==============
'Select the best model and check the residual correlogram

!a = 2
!b = 1
freeze(correlarima{%a}!a!b) arima{%a}!a!b.correl(8)

'Final model
%final ="arimaUNrateAR2MA1"
rename arimaunrate21 {%final}

''==============
'4) Forecasting
''==============
'In sample forecast
'buat dynamic forecast. Kalau nak statik forecast kena letak .fit
smpl %smpl_start %smpl_end
freeze(inforecast) {%final}.forecast(e,g) {%a}f_in
logmsg Fadhilah has completed analysis on in sample forecast and forecast error

'Growth rate comparison
'series UNrate_actual = @pcy({%a})
'series UNrate_in = @pcy({%a}f_in)

series UNrate_actual ={%a}
series UNrate_in = {%a}f_in

smpl 2017 %dat_end
graph g1.line UNrate_actual UNrate_in

'Out-sample forecast
smpl %fcst_start %fcst_end
freeze {%final}.forecast(e,g) {%a}f

smpl %smpl_start %dat_end
series {%a}f = {%a}
smpl @all 
logmsg Fadhilah has completed analysis on out sample forecast

'Growth rate comparison
'series UNrate_arima = @pcy({%a}f)
series UNrate_arima = {%a}f

'=========================
'Monthly  table
'=========================
pagecreate(page=monthly) m %dat_start %fcst_end
string sumvals = "UNrate_actual UNrate_arima"

for %a {sumvals}
copy(c=s) {%file}\{%a} *  'c=s maksudnya copy semuanya s ialah sum. so semua nilai dari gr_arima akan pindah ke page baru
next

'buatkan satu group
group g1

for %b {sumvals}
g1.add ({%b})
next

%x = "arima_monthly"
smpl %dat_end-12 %fcst_end
freeze({%x}) g1.sheet(t) 't maksudnya dalam table

string names = """UNrate-ACTUAL"" ""UNrate-ARIMA"""

!b =0
for %f {names}
!b = !b+1
{%x}(!b+2,1) = %f
next 

{%x}.settextcolor(1) @rgb(0,0,200)
{%x}.setfont(1) b
{%x}.setformat(3,b,4,w) f.2
{%x}.setfont(a) b
{%x}.setjust(a) left
show {%x}

%y = "graph_monthly"
graph {%y}.line g1
{%y}.setelem(1) symbol(filledsquare) linepattern(none)
{%y}.setelem(1) linewidth(3)
{%y}.setelem(2) linewidth(2)
show {%y}

logmsg done{%a} monthly

'=========================
'Quarter table
'=========================
pagecreate(page=qoq) q %dat_start %fcst_end
string sumvals = "UNrate_actual UNrate_arima"

for %a {sumvals}
copy(c=a) {%file}\{%a} *  'tak boleh guna c=s sbb maksudnya copy semuanya s ialah sum. so semua nilai dari gr_arima akan pindah ke page baru. Sepatutnya guna c=a. a ialah averagekan.
next

'buatkan satu group
group g1

for %b {sumvals}
g1.add ({%b})
next

%x = "arima_quarter"
smpl %dat_end-4 %fcst_end
freeze({%x}) g1.sheet(t) 't maksudnya dalam table

string names = """UNrate-ACTUAL"" ""UNrate-ARIMA"""

!b =0
for %f {names}
!b = !b+1
{%x}(!b+2,1) = %f
next 

{%x}.settextcolor(1) @rgb(0,0,200)
{%x}.setfont(1) b
{%x}.setformat(3,b,4,j) f.2
{%x}.setfont(a) b
{%x}.setjust(a) left
show {%x}

%y = "graph_qoq"
graph {%y}.line g1
{%y}.setelem(1) symbol(filledsquare) linepattern(none)
{%y}.setelem(1) linewidth(3)
{%y}.setelem(2) linewidth(2)
show {%y}

logmsg done{%a} quarter

'==========================
'Annual table
'=========================
pagecreate(page=ann) a 2017 %fcst_end
string sumvals = "UNrate_actual UNrate_arima"

for %a {sumvals}
copy(c=a) {%file}\{%a} *  ' a ialah average semua nilai dari gr_arima akan pindah ke page baru
next

'buatkan satu group
group g1

for %b {sumvals}
g1.add ({%b})
next

%x = "arima_annual"
smpl %dat_end-1 %fcst_end '-1 sebab tempoh terakhir data ialah 2019, bila tolak 1 jadi 2018
freeze({%x}) g1.sheet(t) 't maksudnya dalam table

string names = """UNRATE-ACTUAL"" ""UNRATE-ARIMA"""

!b =0
for %f {names}
!b = !b+1
{%x}(!b+2,1) = %f
next 

{%x}.settextcolor(1) @rgb(0,0,200)
{%x}.setfont(1) b
{%x}.setformat(3,b,4,j) f.2
{%x}.setfont(a) b
{%x}.setjust(a) left
show {%x}

%y = "graph_ann"
graph {%y}.line g1
{%y}.setelem(1) symbol(filledsquare) linepattern(none)
{%y}.setelem(1) linewidth(3)
{%y}.setelem(2) linewidth(2)
show {%y}

logmsg done{%a} annual

wfsave(1) {%file}


