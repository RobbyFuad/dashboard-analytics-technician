import streamlit as st
import pandas as pd
import altair as alt
import plotly.express as px
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from pptx.chart.data import CategoryChartData
from pptx.chart.data import ChartData
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

# Feedback extraction
feedback = df['responses.selectedOptions'].str.get_dummies(sep=',')
ordered_feedback = ['Menyelesaikan gangguan dengan cepat',
                    'Datang tepat waktu',
                    'Ramah',
                    'Menjelaskan penyebab gangguan',
                    'Penampilan/ seragam teknisi rapi',
                    'Penyelesaikan gangguan lambat',
                    'Tidak datang tepat waktu',
                    'Kasar dan tidak sopan',
                    'Tidak berseragam/ tidak rapi',
                    'Menjelaskan layanan dengan baik',
                    'Menyelesaikan pemasangan dengan cepat']
feedback = feedback[ordered_feedback]


# FeedbackXRegion
FeedbackXRegion_raw = pd.concat([df['responses.region_new'], feedback], axis=1)
FeedbackXRegion = FeedbackXRegion_raw.groupby(['responses.region_new'], as_index=False).sum()
st.write('Data Dimension: ' + str(FeedbackXRegion.shape[0]) + ' rows and ' + str(FeedbackXRegion.shape[1]) + ' columns.')
st.dataframe(FeedbackXRegion)


# convert long width into long dataframe
selected_option = list(FeedbackXRegion.columns)[1:]
regionXfeedback_long = pd.melt(FeedbackXRegion, id_vars = 'responses.region_new', value_vars = selected_option)
regionXfeedback_long.rename(columns = {'responses.region_new': 'region_fix',
                                       'variable': 'SelectedOption',
                                       'value': 'Count'}, inplace = True)
st.dataframe(regionXfeedback_long)


st.header('Banyaknya Witel di Indonesia')
# Draw banyak_witel chart
banyak_witel_bar = alt.Chart(banyak_witel).mark_bar(size=10).encode(
    y = alt.Y('count:Q', title=None),
    x = alt.X('witel', title=None, sort='-y')
).properties(
    width=1200,
    height=400)

banyak_witel_text = banyak_witel_bar.mark_text(
    align='center',
    baseline='middle',
    #dx=3  # Nudges text to right so it doesn't appear on top of the bar
).encode(
    text='count'
    ).interactive()
(banyak_witel_bar + banyak_witel_text)


st.header('Ringkasan Feedback Teknisi Per Regional')
#Draw FeedbackXRegion
FeedbackXRegion_chart = alt.Chart(regionXfeedback_long).mark_bar().encode(
    x = alt.X('SelectedOption:N', axis=None, sort=alt.EncodingSortField()),
    y = alt.Y('Count:Q', title=None),
    color = alt.Color('SelectedOption:N', sort=alt.EncodingSortField()),
    column = alt.Column('region_fix', title=None)
).properties(
    width=120,
    height=300
).interactive()
FeedbackXRegion_chart


st.header('Ringkasan Feedback Teknisi Per Witel')
# FeedbackXWitel
FeedbackXWitel_raw = pd.concat([df['responses.witel_new'], feedback], axis=1)
FeedbackXWitel = FeedbackXWitel_raw.groupby(['responses.witel_new'], as_index=False).sum()
FeedbackXWitel.rename(columns = {'responses.witel_new': 'witel'}, inplace = True)
FeedbackXWitel_fix = FeedbackXWitel.join(banyak_witel.set_index('witel'), on='witel').sort_values(by=['count'], ascending=False)
st.write(FeedbackXWitel_fix)

# convert long width into long dataframe
FeedbackXwitel_long = pd.melt(FeedbackXWitel_fix, id_vars = 'witel',
                               value_vars = list(FeedbackXWitel_fix.columns)[1:12])
FeedbackXwitel_long.rename(columns = {'variable': 'SelectedOption',
                                       'value': 'Count'}, inplace = True)
st.dataframe(FeedbackXwitel_long)

st.subheader('Top 10 Feedback Teknisi')

#Draw FeedbackXWitel top10
FeedbackXWitel_chart = alt.Chart(FeedbackXwitel_long.head(10)).mark_bar().encode(
    x = alt.X('SelectedOption:N', axis=None, sort=alt.EncodingSortField()),
    y = alt.Y('Count:Q', title=None),
    color = alt.Color('SelectedOption:N', sort=alt.EncodingSortField()),
    column = alt.Column('witel', title=None)
).properties(
    width=120,
    height=300
).interactive()
FeedbackXWitel_chart


st.header('Ringkasan Feedback Teknisi Per STO')
# FeedbackXWitel
FeedbackXSTO_raw = pd.concat([df['responses.sto'], feedback], axis=1)
FeedbackXSTO = FeedbackXSTO_raw.groupby(['responses.sto'], as_index=False).sum()
FeedbackXSTO.rename(columns = {'responses.sto': 'sto'}, inplace = True)
FeedbackXSTO = FeedbackXSTO.join(banyak_sto.set_index('sto'), on='sto').sort_values(by=['count'], ascending=False)
st.write(FeedbackXSTO)


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
    for i in range(1,12):
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
fbreg(FeedbackXRegion, 'Ringkasan Feedback Teknisi Per Region')


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
        for i in range (1,12):
            chart_data.add_series(df.head(10).columns[i], (df.head(10).iloc[:,i]))
    else:
        chart_data.categories = list(df.tail(10).iloc[:,0])
        for i in range (1,12):
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

addSeries(True, FeedbackXWitel_fix, "Ringkasan Feedback Teknisi Per Witel Top 10")
addSeries(False, FeedbackXWitel_fix, "Ringkasan Feedback Teknisi Per Witel Bottom 10")
addSeries(True, FeedbackXSTO, "Ringkasan Feedback Teknisi Per STO Top 10")
addSeries(False, FeedbackXSTO, "Ringkasan Feedback Teknisi Per STO Bottom 10")

last_slide = prs.slides.add_slide(prs.slide_layouts[13])

output = BytesIO()
prs.save(output)
st.download_button(
     label="Download data as PPTX",
     data=output,
     file_name='technician_rate.pptx',
     mime='text/csv'
)

