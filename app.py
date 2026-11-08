"""
Dashboard Kemiskinan dan Ketimpangan di Lima Provinsi (2019-2023).
Sumber data: Badan Pusat Statistik (BPS), periode Maret (Semester 1).
Jalankan dengan:  streamlit run app.py
"""
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =====================================================================
# 1. PENGATURAN (bagian ini boleh kamu edit)
# =====================================================================
FILE_DATA = Path(__file__).parent / "data_kemiskinan_dashboard.xlsx"
TANGGAL_AKSES = "7 November 2026"  # tanggal kamu mengakses data BPS

# Daftar sumber data (judul tabel BPS dan tautannya). Kosongkan "url" bila tidak punya tautan.
SUMBER = [
    {"judul": "Persentase Penduduk Miskin (P0) Menurut Provinsi dan Daerah",
     "url": "https://www.bps.go.id/id/statistics-table/2/MTkyIzI=/persentase-penduduk-miskin--p0--menurut-provinsi-dan-daerah.html"},
    {"judul": "Gini Ratio Menurut Provinsi dan Daerah",
     "url": "https://www.bps.go.id/id/statistics-table/2/OTgjMg==/gini-ratio-menurut-provinsi-dan-daerah.html"},
    {"judul": "Garis Kemiskinan dan Jumlah Penduduk Miskin menurut Provinsi",
     "url": ""},
]

# Tiga temuan utama (judul, isi). Angka berasal dari dataset; penjelasan sebab merujuk RUJUKAN di bawah.
TEMUAN = [
    ("Pandemi menaikkan kemiskinan, tetapi pemulihannya tidak merata",
     "P0 nasional naik dari 9,41% (2019) ke puncak 10,14% (2021), lalu turun ke 9,36% (2023); jumlah penduduk "
     "miskin 2023 (25,90 juta) masih lebih banyak daripada 2019 (25,14 juta). Jawa Tengah dan Jawa Timur sudah "
     "kembali ke level 2019, tetapi P0 DKI Jakarta (+0,97 poin), Jawa Barat (+0,71), dan Bali (+0,46) masih lebih "
     "tinggi. BPS menyebut pandemi COVID-19 sebagai pemicu kenaikan 2020–2021; di Bali, kunjungan wisatawan "
     "mancanegara pada Maret 2021 turun 99,99% dibanding Maret 2020."),
    ("Kemiskinan rendah tidak berarti ketimpangan rendah",
     "DKI Jakarta termasuk provinsi dengan P0 terendah (4,44% pada 2023), tetapi Gini Ratio-nya tertinggi dan naik "
     "setiap tahun: 0,394 (2019) menjadi 0,431 (2023). Jawa Barat juga naik (0,402 menjadi 0,425), dan Jawa Timur "
     "melonjak ke 0,387 pada 2023 meski P0-nya turun tipis. Korelasi Gini–P0 antarprovinsi bernilai negatif "
     "(r = −0,34), yang mencerminkan perbedaan struktur antarprovinsi, bukan bukti bahwa ketimpangan menurunkan "
     "kemiskinan. Kemiskinan sebaiknya dibaca bersama distribusi pendapatan (Bourguignon, 2003)."),
    ("Kesenjangan desa-kota dan garis kemiskinan yang terus naik",
     "Di empat provinsi yang punya data, P0 perdesaan selalu lebih tinggi daripada perkotaan; kesenjangan terbesar "
     "ada di Jawa Timur (13,98% vs 7,50% pada 2023, sekitar 1,9 kali). Garis kemiskinan nominal naik 24–32% selama "
     "2019–2023 (nasional +29,4%), dan komponen makanan menyumbang 74,21% garis kemiskinan nasional pada Maret 2023 "
     "menurut BPS, dengan beras sebagai penyumbang terbesar. Artinya, harga pangan sangat menentukan siapa yang "
     "tergolong miskin, terutama di perdesaan yang porsi berasnya lebih besar (23,73% vs 19,35% di perkotaan; "
     "angka nasional)."),
]

RUJUKAN = """
- BPS. (2021, 15 Juli). *Berita Resmi Statistik: Persentase Penduduk Miskin Maret 2021 Turun Menjadi 10,14 Persen*. \
[bps.go.id](https://www.bps.go.id/pressrelease/2021/07/15/1843/persentase-penduduk-miskin-maret-2021-turun-menjadi-10-14-persen.html)
- Antara. (2022). *BPS: Penduduk miskin turun pada Maret 2022* (memuat penjelasan BPS bahwa kenaikan kemiskinan 2020 \
disebabkan pandemi). [jabar.antaranews.com](https://jabar.antaranews.com/berita/394001/bps-catat-penduduk-miskin-turun-pada-maret-2022)
- Suara Surabaya. (2021, 3 Mei). *BPS: Bali Hanya Terima Kunjungan Tiga Wisman pada Maret 2021* (mengutip Kepala BPS \
Provinsi Bali). [suarasurabaya.net](https://www.suarasurabaya.net/?p=511606)
- Antara. (2023, 17 Juli). *BPS: Garis kemiskinan Maret 2023 sebesar Rp550.458* (komposisi makanan 74,21% dan porsi beras). \
[antaranews.com](https://www.antaranews.com/berita/3638640/bps-garis-kemiskinan-maret-2023-sebesar-rp550458)
- Bourguignon, F. (2003). The growth elasticity of poverty reduction: Explaining heterogeneity across countries and \
time periods. Dalam T. S. Eicher & S. J. Turnovsky (Eds.), *Inequality and Growth: Theory and Policy Implications*. MIT Press.

Angka P0, Gini, dan garis kemiskinan provinsi berasal dari tabel BPS pada bagian sumber data; rujukan di atas hanya
untuk penjelasan sebab dan konteks.
"""

# Palet Okabe-Ito (aman untuk buta warna) + bentuk marker berbeda tiap provinsi,
# sehingga provinsi tidak dibedakan oleh warna saja.
WARNA = {
    "Jawa Timur": "#0072B2",
    "Jawa Tengah": "#E69F00",
    "Jawa Barat": "#009E73",
    "DKI Jakarta": "#D55E00",
    "Bali": "#CC79A7",
}
MARKER = {
    "Jawa Timur": "circle",
    "Jawa Tengah": "square",
    "Jawa Barat": "diamond",
    "DKI Jakarta": "triangle-up",
    "Bali": "star",
}
WARNA_NASIONAL = "#404040"
WARNA_KOTA, WARNA_DESA = "#56B4E9", "#D55E00"

# Nama kolom pada file Excel
C_P0 = "P0 Total (%)"
C_MISKIN = "Jumlah Penduduk Miskin (Ribu Orang)"
C_KOTA = "P0 Perkotaan (%)"
C_DESA = "P0 Perdesaan (%)"
C_GK = "Garis Kemiskinan (Rp/Kapita/Bulan)"
C_GINI = "Gini Ratio (indeks 0-1)"
N_P0 = "P0 Nasional (%)"
N_GK = "Garis Kemiskinan Nasional (Rp/Kapita/Bulan)"
N_GINI = "Gini Ratio Nasional (indeks 0-1)"

CSS = """
<style>
div[data-testid="stMetric"] {
    border: 1px solid #D0D7DE; border-radius: 6px; padding: 12px 16px; background: #FFFFFF;
}
.block-container {padding-top: 2rem;}
</style>
"""


# =====================================================================
# 2. FUNGSI BANTU: data dan format angka
# =====================================================================
@st.cache_data
def muat_data():
    """Membaca dua sheet dari file Excel dan memastikan tipe datanya benar."""
    prov = pd.read_excel(FILE_DATA, sheet_name="Data Provinsi")
    nas = pd.read_excel(FILE_DATA, sheet_name="Data Nasional")
    prov["Tahun"] = prov["Tahun"].astype(int)
    nas["Tahun"] = nas["Tahun"].astype(int)
    return prov, nas


def angka(x, desimal=0):
    """Format angka gaya Indonesia: titik untuk ribuan, koma untuk desimal."""
    s = f"{x:,.{desimal}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def bertanda(x, desimal=0):
    """Angka dengan tanda + atau - di depan (untuk selisih pada KPI)."""
    r = round(x, desimal)
    tanda = "+" if r > 0 else "-" if r < 0 else ""
    return tanda + angka(abs(r), desimal)


def tampil(fig):
    """Menampilkan grafik Plotly; kompatibel dengan berbagai versi Streamlit."""
    cfg = {"displaylogo": False}
    try:
        st.plotly_chart(fig, width="stretch", config=cfg)
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config=cfg)


def periode(a, b):
    return f"Maret {a}" if a == b else f"Maret {a}–{b}"


# =====================================================================
# 3. PERHITUNGAN KPI
# =====================================================================
def hitung_kpi(d):
    """
    d = data provinsi terpilih pada SATU tahun.
    P0 gabungan dihitung tertimbang penduduk (bukan rata-rata sederhana persen):
      populasi implisit = jumlah miskin / (P0 / 100)
      P0 gabungan       = total miskin / total populasi implisit x 100
    """
    populasi = (d[C_MISKIN] / (d[C_P0] / 100)).sum()
    total_miskin = d[C_MISKIN].sum()
    return {
        "p0": total_miskin / populasi * 100,
        "miskin": total_miskin,
        "gk": d[C_GK].mean(),
        "gini": d[C_GINI].mean(),
    }


# =====================================================================
# 4. FUNGSI GRAFIK
# =====================================================================
def gaya(fig, judul, subjudul, x_judul=None, y_judul=None, tinggi=400):
    """Gaya seragam untuk semua grafik: judul, subjudul, sumbu, legenda."""
    fig.update_layout(
        template="plotly_white",
        separators=",.",
        height=tinggi,
        title=dict(text=judul, subtitle=dict(text=subjudul), x=0, xanchor="left"),
        xaxis_title=x_judul,
        yaxis_title=y_judul,
        margin=dict(l=10, r=10, t=85, b=10),
        legend=dict(orientation="h", y=-0.2, x=0, title_text=""),
        font=dict(size=12),
    )
    return fig


def grafik_tren(df, nas, kolom, kol_nas, judul, subjudul, y_judul, fmt, a, b,
                tampil_nas, y_range=None):
    """Line chart tren per provinsi (+ garis nasional putus-putus bila diaktifkan)."""
    fig = go.Figure()
    for p in WARNA:
        d = df[df["Provinsi"] == p].sort_values("Tahun")
        if d.empty:
            continue
        fig.add_trace(go.Scatter(
            x=d["Tahun"], y=d[kolom], name=p, mode="lines+markers",
            line=dict(color=WARNA[p], width=2.5),
            marker=dict(symbol=MARKER[p], size=9, color=WARNA[p]),
            hovertemplate=f"<b>{p}</b><br>Tahun %{{x}}<br>{y_judul}: %{{y:{fmt}}}<extra></extra>",
        ))
    if tampil_nas:
        fig.add_trace(go.Scatter(
            x=nas["Tahun"], y=nas[kol_nas], name="Nasional (Indonesia)", mode="lines+markers",
            line=dict(color=WARNA_NASIONAL, width=2, dash="dash"),
            marker=dict(symbol="circle-open", size=8, color=WARNA_NASIONAL),
            hovertemplate=f"<b>Nasional</b><br>Tahun %{{x}}<br>{y_judul}: %{{y:{fmt}}}<extra></extra>",
        ))
    gaya(fig, judul, subjudul, "Tahun (data Maret)", y_judul)
    fig.update_xaxes(dtick=1, tickformat="d", range=[a - 0.4, b + 0.4])
    if y_range:
        fig.update_yaxes(range=y_range)
    else:
        fig.update_yaxes(rangemode="tozero")
    return fig


def grafik_desa_kota(d, b):
    """Grouped bar P0 perkotaan vs perdesaan pada tahun akhir. Sumbu Y mulai dari 0."""
    fig = go.Figure()
    for kol, nama, warna, pola in [(C_KOTA, "Perkotaan", WARNA_KOTA, ""), (C_DESA, "Perdesaan", WARNA_DESA, "/")]:
        fig.add_trace(go.Bar(
            x=d["Provinsi"], y=d[kol], name=nama,
            marker=dict(color=warna, pattern=dict(shape=pola)),
            text=[angka(v, 2) if pd.notna(v) else "" for v in d[kol]], textposition="outside",
            hovertemplate=f"<b>%{{x}}</b><br>{nama}: %{{y:.2f}}%<extra></extra>",
        ))
    # Data kosong ditandai teks, bukan digambar sebagai nol
    for _, r in d[d[C_DESA].isna()].iterrows():
        fig.add_annotation(x=r["Provinsi"], y=r[C_KOTA], yshift=38, showarrow=False,
                           text="Perdesaan:<br>tidak tersedia", font=dict(size=10, color="#555555"))
    gaya(fig, "Penduduk miskin di perkotaan dan perdesaan", f"Persen penduduk, {periode(b, b)}",
         "Provinsi", "Persentase penduduk miskin (%)")
    fig.update_layout(barmode="group")
    fig.update_yaxes(range=[0, float(d[[C_KOTA, C_DESA]].max().max()) * 1.35])
    return fig


def grafik_scatter(d, a, b):
    """Scatter Gini (x) vs P0 (y); satu titik = satu provinsi-tahun; garis tren OLS."""
    pakai_tren = len(d) >= 3 and d[C_GINI].nunique() > 1
    fig = px.scatter(
        d, x=C_GINI, y=C_P0, color="Provinsi", symbol="Provinsi",
        color_discrete_map=WARNA, symbol_map=MARKER, hover_name="Provinsi",
        hover_data={"Tahun": True, C_GINI: ":.3f", C_P0: ":.2f", C_MISKIN: ":,.0f", "Provinsi": False},
        trendline="ols" if pakai_tren else None, trendline_scope="overall",
        trendline_color_override=WARNA_NASIONAL,
    )
    fig.update_traces(marker=dict(size=12, opacity=0.9, line=dict(width=1, color="white")),
                      selector=dict(mode="markers"))
    for tr in fig.data:
        if tr.mode == "lines":
            tr.name, tr.showlegend = "Garis tren OLS (semua titik)", True
    gaya(fig, "Hubungan Gini Ratio dan persentase penduduk miskin",
         f"Satu titik = satu provinsi-tahun, {periode(a, b)}",
         "Gini Ratio (indeks 0–1)", "Persentase penduduk miskin (%)")
    fig.update_yaxes(rangemode="tozero")
    return fig


def grafik_garis(d, nas_gk, b):
    """Bar horizontal garis kemiskinan tahun akhir, terurut, dengan penanda nasional."""
    d = d.sort_values(C_GK)
    fig = go.Figure(go.Bar(
        y=d["Provinsi"], x=d[C_GK], orientation="h",
        marker=dict(color=[WARNA[p] for p in d["Provinsi"]]),
        text=[f"Rp {angka(v)}" for v in d[C_GK]], textposition="outside",
        hovertemplate="<b>%{y}</b><br>Rp %{x:,.0f} per kapita per bulan<extra></extra>",
    ))
    fig.add_vline(x=nas_gk, line_dash="dash", line_color=WARNA_NASIONAL,
                  annotation_text=f"Nasional: Rp {angka(nas_gk)}", annotation_position="top")
    gaya(fig, "Garis kemiskinan per provinsi",
         f"Rupiah per kapita per bulan, {periode(b, b)} (nominal, belum disesuaikan inflasi)",
         "Garis kemiskinan (Rp per kapita per bulan)", None, tinggi=380)
    fig.update_xaxes(range=[0, max(float(d[C_GK].max()), nas_gk) * 1.25], tickformat=",.0f")
    return fig


# =====================================================================
# 5. HALAMAN UTAMA
# =====================================================================
def main():
    st.set_page_config(page_title="Dashboard Kemiskinan dan Ketimpangan", page_icon="📊", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    prov, nas = muat_data()

    # ---- Filter (sidebar) ----
    st.sidebar.header("Filter")
    t_min, t_max = int(prov["Tahun"].min()), int(prov["Tahun"].max())
    a, b = st.sidebar.slider("Rentang tahun (data Maret)", t_min, t_max, (t_min, t_max), step=1)
    daftar = [p for p in WARNA if p in set(prov["Provinsi"])]
    pilihan = st.sidebar.multiselect("Provinsi", daftar, default=daftar)
    tampil_nas = st.sidebar.toggle("Tampilkan pembanding nasional", value=True)
    st.sidebar.caption("Semua indikator dan grafik mengikuti filter di atas.")

    st.title("Kemiskinan dan Ketimpangan di Lima Provinsi, 2019–2023")
    st.caption("Persentase penduduk miskin (P0), garis kemiskinan, dan Gini Ratio. Sumber: BPS, periode Maret setiap tahun.")

    if not pilihan:
        st.warning("Belum ada provinsi yang dipilih. Pilih minimal satu provinsi pada menu Provinsi di sebelah kiri.")
        st.stop()

    df = prov[prov["Tahun"].between(a, b) & prov["Provinsi"].isin(pilihan)]
    nas_f = nas[nas["Tahun"].between(a, b)]
    akhir, awal = df[df["Tahun"] == b], df[df["Tahun"] == a]
    k_akhir, k_awal = hitung_kpi(akhir), hitung_kpi(awal)
    banding = a != b  # selisih hanya ditampilkan jika tahun awal berbeda dari tahun akhir
    ket = f"dibanding {a}"

    # ---- KPI ----
    st.subheader(f"Indikator utama, Maret {b}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("P0 gabungan (%)", f"{angka(k_akhir['p0'], 2)}%",
              f"{bertanda(k_akhir['p0'] - k_awal['p0'], 2)} poin vs {a}" if banding else None,
              delta_color="off",
              help="Total penduduk miskin dibagi total penduduk provinsi terpilih (tertimbang penduduk). "
                   "Penduduk diturunkan dari jumlah miskin dibagi P0, sehingga provinsi besar berbobot lebih besar.")
    c2.metric("Penduduk miskin", angka(k_akhir["miskin"], 0),
              f"{bertanda(k_akhir['miskin'] - k_awal['miskin'], 0)} ribu vs {a}" if banding else None,
              delta_color="off", help="Jumlah penduduk miskin seluruh provinsi terpilih, dalam ribu orang.")
    c3.metric("Garis kemiskinan", f"Rp {angka(k_akhir['gk'], 0)}",
              f"{bertanda(k_akhir['gk'] - k_awal['gk'], 0)} Rp vs {a}" if banding else None,
              delta_color="off",
              help="Rata-rata sederhana antarprovinsi, Rp per kapita per bulan. Nilai nominal, belum disesuaikan inflasi.")
    c4.metric("Gini Ratio", angka(k_akhir["gini"], 3),
              f"{bertanda(k_akhir['gini'] - k_awal['gini'], 3)} vs {a}" if banding else None,
              delta_color="off",
              help="Rata-rata sederhana antarprovinsi, indeks 0 sampai 1. Makin tinggi, makin timpang.")
    st.caption("Satuan: P0 dalam persen (tertimbang penduduk); penduduk miskin dalam ribu orang; garis kemiskinan dalam "
               "Rp per kapita per bulan (rata-rata sederhana antarprovinsi, nominal); Gini Ratio indeks 0–1 "
               "(rata-rata sederhana antarprovinsi). Selisih dibandingkan dengan tahun awal rentang.")
    if not banding:
        st.caption("Selisih tidak ditampilkan karena tahun awal dan tahun akhir sama. Geser slider untuk membandingkan.")

    # ---- Grafik ----
    st.subheader("Visualisasi")
    kiri, kanan = st.columns(2)
    with kiri:
        tampil(grafik_tren(df, nas_f, C_P0, N_P0, "Tren persentase penduduk miskin (P0) per provinsi",
                           f"Persen penduduk, {periode(a, b)}", "Persentase penduduk miskin (%)", ".2f",
                           a, b, tampil_nas))
    with kanan:
        tampil(grafik_tren(df, nas_f, C_GINI, N_GINI, "Tren Gini Ratio per provinsi",
                           f"Indeks 0–1, makin tinggi makin timpang, {periode(a, b)}", "Gini Ratio (indeks 0–1)",
                           ".3f", a, b, tampil_nas, y_range=[0.30, 0.50]))
        st.caption("Catatan skala: sumbu Y dimulai dari 0,30 (bukan 0) agar selisih antarprovinsi terbaca; "
                   "secara teori Gini berkisar 0 sampai 1.")

    kiri, kanan = st.columns(2)
    with kiri:
        tampil(grafik_desa_kota(akhir, b))
    with kanan:
        tampil(grafik_scatter(df, a, b))
        if len(df) >= 3 and df[C_GINI].nunique() > 1:
            r = df[C_GINI].corr(df[C_P0])
            st.caption(f"Korelasi Pearson (r) = {angka(r, 2)}, n = {len(df)} pengamatan provinsi-tahun. "
                       f"Korelasi bukan kausalitas; data hanya {len(pilihan)} provinsi, sehingga perbedaan "
                       "antarprovinsi sangat memengaruhi hasil.")
        else:
            st.caption("Korelasi dan garis tren muncul jika ada minimal tiga pengamatan. "
                       "Perluas rentang tahun atau tambah provinsi.")

    nas_gk = float(nas.loc[nas["Tahun"] == b, N_GK].iloc[0])
    tampil(grafik_garis(akhir, nas_gk, b))

    # ---- Temuan utama ----
    st.subheader("Temuan utama")
    st.caption("Berlaku untuk data Maret 2019–2023 pada lima provinsi; teks ini tidak berubah mengikuti filter.")
    for kolom_t, (judul_t, isi_t) in zip(st.columns(3), TEMUAN):
        with kolom_t.container(border=True):
            st.markdown(f"**{judul_t}**")
            st.write(isi_t)
    with st.expander("Rujukan untuk temuan utama"):
        st.markdown(RUJUKAN)

    # ---- Tabel data ----
    with st.expander("Lihat dan unduh data yang sedang ditampilkan"):
        tabel = df.copy()
        tabel[C_DESA] = tabel[C_DESA].map(lambda v: "tidak tersedia" if pd.isna(v) else angka(v, 2))
        st.dataframe(tabel, hide_index=True)
        st.download_button("Unduh data terfilter (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                           "data_terfilter.csv", "text/csv")

    # ---- Sumber dan catatan ----
    st.divider()
    baris = []
    for src in SUMBER:
        if src["url"]:
            url = src["url"].replace("(", "%28").replace(")", "%29").replace(" ", "%20")
            baris.append(f"- [{src['judul']}]({url})")
        else:
            baris.append(f"- {src['judul']}")
    st.markdown("**Sumber data:** Badan Pusat Statistik (BPS), periode Maret (Semester 1) 2019–2023. "
                f"Diakses: {TANGGAL_AKSES}.\n\n" + "\n".join(baris))
    st.caption(
        "Catatan: P0 perdesaan DKI Jakarta tidak tersedia di sumber, jadi dikosongkan dan tidak dihitung "
        "(bukan nol). Garis kemiskinan berupa nilai nominal. Data hanya periode Maret, bukan rata-rata setahun."
    )


if __name__ == "__main__":
    main()
