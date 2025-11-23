import pandas as pd
from fuzzywuzzy import fuzz
from tkinter import Tk, Label, Button, filedialog, messagebox, Text, OptionMenu, StringVar
import threading

class FuzzyMatcherGUI:
    def __init__(self, master):
        self.master = master
        master.title("Fuzzy Matching Excel - IDSBR Best Match Only")

        Label(master, text="File Excel 1 (df1):").pack()
        self.file1_btn = Button(master, text="Pilih File", command=self.load_file1)
        self.file1_btn.pack()

        self.idsbr_label = Label(master, text="Pilih kolom IDSBR:")
        self.idsbr_var = StringVar(master)
        self.idsbr_dropdown = None

        Label(master, text="File Excel 2 (df2):").pack()
        self.file2_btn = Button(master, text="Pilih File", command=self.load_file2)
        self.file2_btn.pack()

        self.run_btn = Button(master, text="Mulai Pencocokan", command=self.run_match_thread)
        self.run_btn.pack(pady=10)

        self.text_output = Text(master, height=20, width=100)
        self.text_output.pack()

        self.file1 = None
        self.file2 = None
        self.df1_headers = []

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        self.log(f"📂 File 1 dipilih: {self.file1}")
        try:
            df1 = pd.read_excel(self.file1, nrows=1)
            self.df1_headers = list(df1.columns)
            self.show_idsbr_dropdown()
        except Exception as e:
            self.log(f"❌ Gagal membaca header dari File 1: {e}")

    def show_idsbr_dropdown(self):
        if self.idsbr_dropdown:
            self.idsbr_dropdown.destroy()

        self.idsbr_label.pack()
        self.idsbr_var.set(self.df1_headers[0])
        self.idsbr_dropdown = OptionMenu(self.master, self.idsbr_var, *self.df1_headers)
        self.idsbr_dropdown.pack()

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        self.log(f"📂 File 2 dipilih: {self.file2}")

    def run_match_thread(self):
        t = threading.Thread(target=self.run_match)
        t.start()

    def run_match(self):
        if not self.file1 or not self.file2:
            messagebox.showerror("Error", "Silakan pilih kedua file terlebih dahulu.")
            return

        try:
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)
        except Exception as e:
            self.log(f"❌ Error membaca file: {e}")
            return

        idsbr_col = self.idsbr_var.get()
        required_cols = [idsbr_col, 'nama_usaha', 'alamat_usaha', 'kode_wilayah']
        try:
            df1 = df1[required_cols].dropna()
            df2 = df2[['nama_usaha', 'alamat_usaha', 'kode_wilayah']].dropna()
        except Exception as e:
            self.log(f"❌ Kolom wajib tidak ditemukan: {e}")
            return

        best_matches = []
        wilayah_set = set(df1['kode_wilayah']).intersection(df2['kode_wilayah'])

        for wilayah in wilayah_set:
            df1_grp = df1[df1['kode_wilayah'] == wilayah]
            df2_grp = df2[df2['kode_wilayah'] == wilayah]

            self.log(f"🔍 Mencocokkan wilayah {wilayah}: {len(df1_grp)} x {len(df2_grp)} baris")

            for _, r1 in df1_grp.iterrows():
                best_match = None
                best_score = -1

                for _, r2 in df2_grp.iterrows():
                    name_score = fuzz.token_sort_ratio(str(r1['nama_usaha']), str(r2['nama_usaha']))
                    addr_score = fuzz.token_sort_ratio(str(r1['alamat_usaha']), str(r2['alamat_usaha']))
                    avg_score = (name_score + addr_score) / 2

                    if avg_score > best_score:
                        best_score = avg_score
                        best_match = {
                            'IDSBR': r1[idsbr_col],
                            'df1_nama': r1['nama_usaha'],
                            'df2_nama': r2['nama_usaha'],
                            'df1_alamat': r1['alamat_usaha'],
                            'df2_alamat': r2['alamat_usaha'],
                            'kode_wilayah': wilayah,
                            'nama_score': name_score,
                            'alamat_score': addr_score,
                            'avg_score': avg_score
                        }

                if best_match:
                    best_matches.append(best_match)

        if best_matches:
            match_df = pd.DataFrame(best_matches)
            output_file = "fuzzy_best_matches.xlsx"
            match_df.to_excel(output_file, index=False)
            self.log(f"✅ Pencocokan selesai. Total IDSBR: {len(best_matches)}")
            self.log(f"📁 Hasil disimpan: {output_file}")
        else:
            self.log("⚠️ Tidak ditemukan pasangan yang cocok.")

    def log(self, message):
        self.text_output.insert("end", message + "\n")
        self.text_output.see("end")

if __name__ == "__main__":
    root = Tk()
    app = FuzzyMatcherGUI(root)
    root.mainloop()
