from doctest import DocFileTest
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor

st.title("Technician Rating")  # add a title
st.header('Latar Belakang')
st.markdown('<div style="text-align: justify"> Kepuasan pelanggan dapat diukur salah satunya dari aspek kepuasan pelayanan terhadap pelanggan. Jika pelayanan yang diterima oleh pelanggan tidak/memenuhi harapan mereka, maka pelanggan akan memberikan feedback sesuai apa yang mereka peroleh. Data rating teknisi diperlukan untuk mengevaluasi teknisi apakah mereka telah menjalankan tugas dengan baik atau tidak. </div>', unsafe_allow_html=True)
st.markdown('''
Alur untuk mengisi survey Data Teknisi:
* Isi FAQ.
* Melaporkan gangguan.
* Menjadwalkan kapan untuk perbaikan perangkat.
* Repair Request (Riwayat Tracking Teknisi berangkat hingga problem solved).
* Setelah problem solved, pelanggan mengisi survey Data Teknisi (rating, alasan memberikan rating tersebut, regional, witel, id teknisi, data diri pengguna, dan lain-lain).
''')

uploaded_file = st.file_uploader("Choose a file")

#@st.cache
df = pd.read_csv('technicianrate.csv', sep=';')
df.drop(df.columns[32:],axis=1,inplace=True)
cols = [8,21,22]
df.drop(df.columns[cols],axis=1,inplace=True)
df['tanggal']=pd.to_datetime(df['tanggal'], format="%d/%m/%Y")
df['tanggal'] = df['tanggal'].dt.date

st.sidebar.header('Input Feature')
selected_cols = st.sidebar.selectbox('kolom', df.columns)

witel = df['responses.witel_new'].unique()
selected_witel = st.sidebar.multiselect('Witel', witel, witel[:1])

st.sidebar.header('Input Feature')
df_selected_feature = df['responses.witel_new'].isin(selected_witel)
#df_selected_feature = df[(df.columns.isin(selected_cols)) & (df['responses.witel_new'].unique().isin(selected_witel))]

st.header('Dataset yang telah dipilih')
st.write('Data Dimension: ' + str(df[df_selected_feature].shape[0]) + ' rows and ' + str(df[df_selected_feature].shape[1]) + ' columns.')
st.dataframe(df[df_selected_feature])

banyak_witel = df['responses.witel_new'].value_counts().rename_axis('witel').reset_index(name='count').sort_values(by=['witel'])
banyak_sto = df['responses.sto'].value_counts().rename_axis('sto').reset_index(name='count').sort_values(by=['sto'])

# Rating Extraction
rating = pd.get_dummies(df['responses.rate'], columns=['responses.rate'], prefix='rating', prefix_sep=' ')

# RateXRegion
RateXRegion_raw = pd.concat([df['responses.region_new'], rating], axis=1)
RateXRegion = RateXRegion_raw.groupby(['responses.region_new'], as_index=False).sum()
# RateXWitel
RateXWitel_raw = pd.concat([df['responses.witel_new'], rating], axis=1)
RateXWitel = RateXWitel_raw.groupby(['responses.witel_new'], as_index=False).sum()
RateXWitel.rename(columns = {'responses.witel_new': 'witel'}, inplace = True)
RateXWitel_fix = RateXWitel.join(banyak_witel.set_index('witel'), on='witel').sort_values(by=['count'], ascending=False)
# RateXSTO
RateXSTO_raw = pd.concat([df['responses.sto'], rating], axis=1)
RateXSTO = RateXSTO_raw.groupby(['responses.sto'], as_index=False).sum()
RateXSTO.rename(columns = {'responses.sto': 'sto'}, inplace = True)
RateXSTO = RateXSTO.join(banyak_sto.set_index('sto'), on='sto').sort_values(by=['count'], ascending=False)
RateXSTO = RateXSTO.astype({
                    'rating 1':int,
                    'rating 2':int,
                    'rating 3':int,
                    'rating 4':int,
                    'rating 5':int,
})
# Feedback extraction
feedback = df['responses.selectedOptions'].str.get_dummies(sep=',')
ordered_feedback = pd.DataFrame(feedback.sum(axis=0).sort_values(ascending=False)).index.tolist() #order column based on their value
feedback = feedback[ordered_feedback]

# FeedbackXRegion
FeedbackXRegion_raw = pd.concat([df['responses.region_new'], feedback], axis=1)
FeedbackXRegion = FeedbackXRegion_raw.groupby(['responses.region_new'], as_index=False).sum()
# FeedbackXWitel
FeedbackXWitel_raw = pd.concat([df['responses.witel_new'], feedback], axis=1)
FeedbackXWitel = FeedbackXWitel_raw.groupby(['responses.witel_new'], as_index=False).sum()
FeedbackXWitel.rename(columns = {'responses.witel_new': 'witel'}, inplace = True)
FeedbackXWitel_fix = FeedbackXWitel.join(banyak_witel.set_index('witel'), on='witel').sort_values(by=['count'], ascending=False)
# FeedbackXSTO
FeedbackXSTO_raw = pd.concat([df['responses.sto'], feedback], axis=1)
FeedbackXSTO = FeedbackXSTO_raw.groupby(['responses.sto'], as_index=False).sum()
FeedbackXSTO.rename(columns = {'responses.sto': 'sto'}, inplace = True)
FeedbackXSTO = FeedbackXSTO.join(banyak_sto.set_index('sto'), on='sto').sort_values(by=['count'], ascending=False)

# extract date
bulan = pd.concat([df['tanggal'], rating], axis=1)
# rateXbulan = bulan.groupby(['tanggal'], as_index=False).sum()
# rateXbulan['tanggal'] = rateXbulan['tanggal'].dt.strftime('%m/%Y')
# rateXbulan_fix = rateXbulan.groupby(['tanggal'], as_index=False).sum()

# function to call graph
def plotlygraph(str_data, df, str_mode):
    if str_data=='head':
        fig = px.bar(df.head(10),
                     x=df.head(10).columns[0],
                     y=df.head(10).columns[1:-1],
                     barmode=str_mode
                     #title="Wide-Form Input",
                    )
    elif str_data=='tail':
        fig = px.bar(df.tail(10),
                     x=df.tail(10).columns[0],
                     y=df.tail(10).columns[1:-1],
                     barmode=str_mode
                    )
    else:
        fig = px.bar(df,
                     x=df.columns[0],
                     y=df.columns[1:],
                     barmode=str_mode
                    )
    st.plotly_chart(fig, use_container_width=True)   
    return

####################################################################################################################################################
st.header('Ringkasan Rating Teknisi Per Region')
plotlygraph(None, RateXRegion, 'group')

st.header('Ringkasan Rating Teknisi Per Witel')
st.subheader('Top 10 Rating')
plotlygraph('head', RateXWitel_fix, 'group')
st.subheader('Bottom 10 Rating')
plotlygraph('tail', RateXWitel_fix, 'group')

st.header('Ringkasan Rating Teknisi Per STO')
# st.write(RateXSTO)
st.subheader('Top 10 Rating')
plotlygraph('head', RateXSTO, 'group')
st.subheader('Bottom 10 Rating')
plotlygraph('tail', RateXSTO, 'group')

####
st.header('Ringkasan Feedback Teknisi Per Region')
plotlygraph(None, FeedbackXRegion, 'group')

st.header('Ringkasan Feedback Teknisi Per Witel')
st.subheader('Top 10 Feedback')
plotlygraph('head', FeedbackXWitel_fix, 'group')
st.subheader('Bottom 10 Feedback')
plotlygraph('tail', FeedbackXWitel_fix, 'group')

st.header('Ringkasan Feedback Teknisi Per STO')
# st.write(FeedbackXSTO)
st.subheader('Top 10 Feedback')
plotlygraph('head', FeedbackXSTO, 'group')
st.subheader('Bottom 10 Feedback')
plotlygraph('tail', FeedbackXSTO, 'group')

st.header('Ringkasan Rating Teknisi Per Bulan')
st.write(feedback)


####################################################################################################################################################

# MAKE PPTX FILE
prs = Presentation('slidemstr.pptx')
# prs.slide_width = Inches(16)
# prs.slide_height = Inches(9)
slide_reg = prs.slide_layouts[11]

# Make fbXreg slide
def fbreg(df, str_title):
    slide = prs.slides.add_slide(slide_reg)
    title = slide.shapes.title
    # .left .top .width .hight
    title.width.top = Inches(3), Inches(3)
    title.text = str_title
    # title1.text_frame.paragraphs[0].font.name = "Arial"
    chart_data = ChartData()
    chart_data.categories = list(df.iloc[:,0])
    for i in range(1,len(df.columns)):
        chart_data.add_series(df.columns[i], (df.iloc[:,i]))
    
    x, y, cx, cy = Inches(1), Inches(1), Inches(12), Inches(6)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart
    
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels

    data_labels.font.size = Pt(8)
    data_labels.font.color.rgb = RGBColor(0,0,0)
    data_labels.position = XL_LABEL_POSITION.INSIDE_END
    
    chart.category_axis.tick_labels.font.size = Pt(11)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(11)
    chart.legend.font.name = 'Arial'
    return


def addSeries(head, df, str_title):
    slide = prs.slides.add_slide(slide_reg)
    title = slide.shapes.title
    # .left .top .width .hight
    title.width.top = Inches(3), Inches(3)
    # title3.top = Inches(3)
    title.text = str_title
    chart_data = ChartData()
    if head:
        chart_data.categories = list(df.head(10).iloc[:,0])
        for i in range (1,len(df.columns)-1):
            chart_data.add_series(df.head(10).columns[i], (df.head(10).iloc[:,i]))
    else:
        chart_data.categories = list(df.tail(10).iloc[:,0])
        for i in range (1,len(df.columns)-1):
            chart_data.add_series(df.tail(10).columns[i], (df.tail(10).iloc[:,i]))

    x, y, cx, cy = Inches(1), Inches(1), Inches(12), Inches(6)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart

    chart.category_axis.tick_labels.font.size = Pt(11)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(11)
    chart.legend.font.name = 'Arial'
    return


fbreg(RateXRegion, 'Ringkasan Rating Teknisi Per Region')
addSeries(True, RateXWitel_fix, "Ringkasan Rating Teknisi Per Witel Top 10")
addSeries(False, RateXWitel_fix, "Ringkasan Rating Teknisi Per Witel Bottom 10")
addSeries(True, RateXSTO, "Ringkasan Rating Teknisi Per STO Top 10")
addSeries(False, RateXSTO, "Ringkasan Rating Teknisi Per STO Bottom 10")
fbreg(FeedbackXRegion, 'Ringkasan Feedback Teknisi Per Region')
addSeries(True, FeedbackXWitel_fix, "Ringkasan Feedback Teknisi Per Witel Top 10")
addSeries(False, FeedbackXWitel_fix, "Ringkasan Feedback Teknisi Per Witel Bottom 10")
addSeries(True, FeedbackXSTO, "Ringkasan Feedback Teknisi Per STO Top 10")
addSeries(False, FeedbackXSTO, "Ringkasan Feedback Teknisi Per STO Bottom 10")

last_slide = prs.slides.add_slide(prs.slide_layouts[13])

output = BytesIO()
prs.save(output)
st.download_button(
     label="Generate PPTX",
     data=output,
     file_name='technician_rate.pptx',
     mime='text/csv'
)

