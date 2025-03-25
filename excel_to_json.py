import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, 
                            QHBoxLayout, QFileDialog, QListWidget, QWidget, QCheckBox,
                            QMessageBox, QProgressBar, QFrame, QGroupBox, QDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QFont

class AboutDialog(QDialog):
    """Hakkında bilgilerini gösteren iletişim kutusu"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Hakkında")
        self.setFixedSize(300, 150)
        
        layout = QVBoxLayout()
        
        # Uygulama adı
        app_name = QLabel("Excel → JSON Dönüştürücü")
        app_name.setFont(QFont("", 14, QFont.Bold))
        app_name.setAlignment(Qt.AlignCenter)
        layout.addWidget(app_name)
        
        # Geliştirici bilgisi
        dev_label = QLabel("Geliştirici: Mehmet ÇİMEN")
        dev_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(dev_label)
        
        # Website
        website = QLabel('<a href="https://mehmetc.dev">https://mehmetc.dev</a>')
        website.setOpenExternalLinks(True)
        website.setAlignment(Qt.AlignCenter)
        layout.addWidget(website)
        
        # Sürüm bilgisi
        version = QLabel("Sürüm: 1.0")
        version.setAlignment(Qt.AlignCenter)
        layout.addWidget(version)
        
        # Kapat butonu
        close_btn = QPushButton("Kapat")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)

class ConversionThread(QThread):
    """Excel'den JSON'a dönüşüm işlemini ayrı bir thread'de yapar"""
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    
    def __init__(self, excel_file, selected_sheets, output_file):
        super().__init__()
        self.excel_file = excel_file
        self.selected_sheets = selected_sheets
        self.output_file = output_file
        
    def run(self):
        try:
            # Excel dosyasını oku
            xl = pd.ExcelFile(self.excel_file)
            
            # Seçilen sayfalar varsa onları kullan, yoksa tüm sayfaları kullan
            sheets_to_convert = self.selected_sheets if self.selected_sheets else xl.sheet_names
            
            # JSON verisi için boş sözlük
            excel_data = {}
            
            # Toplam sayfa sayısı
            total_sheets = len(sheets_to_convert)
            
            # Her sayfayı dönüştür
            for idx, sheet_name in enumerate(sheets_to_convert):
                # İlerleme durumunu bildir
                progress = int((idx / total_sheets) * 100)
                self.progress_signal.emit(progress)
                
                # Sayfayı oku
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                
                # NaN değerleri None olarak değiştir
                df = df.where(pd.notnull(df), None)
                
                # DataFrame'i dict listesine dönüştür
                records = df.to_dict(orient='records')
                
                # Sheet adını anahtar olarak kullan
                excel_data[sheet_name] = records
            
            # JSON dosyasına yaz
            with open(self.output_file, 'w', encoding='utf-8') as json_file:
                json.dump(excel_data, json_file, ensure_ascii=False, indent=4)
            
            # %100 ilerleme
            self.progress_signal.emit(100)
            
            # İşlem tamamlandı sinyali gönder
            self.finished_signal.emit(self.output_file)
            
        except Exception as e:
            # Hata oluşursa sinyal gönder
            self.error_signal.emit(str(e))

class ExcelToJsonApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
        # Değişkenler
        self.excel_file = ""
        self.sheets = []
        self.selected_sheets = []
        
    def init_ui(self):
        # Ana pencere özellikleri
        self.setWindowTitle('Excel → JSON Dönüştürücü')
        self.setGeometry(300, 300, 600, 500)
        self.setMinimumSize(500, 400)
        
        # Ana widget ve layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # Başlık
        title_label = QLabel('Excel\'den JSON\'a Dönüştürme Aracı')
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # Dosya seçimi bölümü
        file_group = QGroupBox("Excel Dosyası")
        file_layout = QVBoxLayout()
        file_group.setLayout(file_layout)
        
        # Dosya seçim satırı
        file_row = QHBoxLayout()
        self.file_label = QLabel('Dosya seçilmedi')
        file_btn = QPushButton('Excel Dosyası Seç')
        file_btn.clicked.connect(self.select_file)
        file_row.addWidget(self.file_label, 1)
        file_row.addWidget(file_btn, 0)
        file_layout.addLayout(file_row)
        
        main_layout.addWidget(file_group)
        
        # Sayfa seçimi bölümü
        self.sheets_group = QGroupBox("Dönüştürülecek Sayfalar")
        self.sheets_group.setEnabled(False)
        sheets_layout = QVBoxLayout()
        self.sheets_group.setLayout(sheets_layout)
        
        # Tümünü seç/hiçbirini seçme
        select_row = QHBoxLayout()
        select_all_btn = QPushButton('Tümünü Seç')
        select_all_btn.clicked.connect(self.select_all_sheets)
        select_none_btn = QPushButton('Tümünü Kaldır')
        select_none_btn.clicked.connect(self.select_no_sheets)
        select_row.addWidget(select_all_btn)
        select_row.addWidget(select_none_btn)
        sheets_layout.addLayout(select_row)
        
        # Sayfa listesi
        self.sheets_list = QListWidget()
        self.sheets_list.setSelectionMode(QListWidget.MultiSelection)
        sheets_layout.addWidget(self.sheets_list)
        
        main_layout.addWidget(self.sheets_group)
        
        # Çıktı dosyası bölümü
        output_group = QGroupBox("JSON Çıktı")
        output_layout = QVBoxLayout()
        output_group.setLayout(output_layout)
        
        output_row = QHBoxLayout()
        self.output_label = QLabel('Henüz seçilmedi')
        output_btn = QPushButton('Çıktı Konumunu Seç')
        output_btn.clicked.connect(self.select_output)
        output_row.addWidget(self.output_label, 1)
        output_row.addWidget(output_btn, 0)
        output_layout.addLayout(output_row)
        
        main_layout.addWidget(output_group)
        
        # İlerleme çubuğu
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Alt bölüm (Dönüştür butonu ve Hakkında butonu)
        buttons_layout = QHBoxLayout()
        
        # Dönüştür butonu
        convert_btn = QPushButton('Dönüştür')
        convert_btn.setMinimumHeight(40)
        convert_btn_font = QFont()
        convert_btn_font.setBold(True)
        convert_btn.setFont(convert_btn_font)
        convert_btn.clicked.connect(self.convert)
        buttons_layout.addWidget(convert_btn)
        
        # Hakkında butonu
        about_btn = QPushButton('Hakkında')
        about_btn.clicked.connect(self.show_about)
        buttons_layout.addWidget(about_btn)
        
        main_layout.addLayout(buttons_layout)
        
        # Durum çubuğu
        self.statusBar().showMessage('Hazır')
        
    def show_about(self):
        """Hakkında iletişim kutusunu göster"""
        about_dialog = AboutDialog(self)
        about_dialog.exec_()
        
    def select_file(self):
        """Excel dosyası seçme işlemi"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'Excel Dosyası Seç', '', 'Excel Dosyaları (*.xlsx *.xls);;Tüm Dosyalar (*)')
        
        if file_path:
            self.excel_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.load_sheets()
            
            # Varsayılan çıktı dosyası adını ayarla
            base_name = os.path.splitext(self.excel_file)[0]
            self.output_label.setText(f"{base_name}.json")
    
    def load_sheets(self):
        """Excel dosyasındaki sayfaları yükler"""
        try:
            xl = pd.ExcelFile(self.excel_file)
            self.sheets = xl.sheet_names
            
            # Mevcut listeyi temizle
            self.sheets_list.clear()
            
            # Sayfaları listeye ekle
            for sheet in self.sheets:
                self.sheets_list.addItem(sheet)
            
            # Varsayılan olarak tüm sayfaları seç
            for i in range(self.sheets_list.count()):
                self.sheets_list.item(i).setSelected(True)
            
            # Sayfa seçim grubunu etkinleştir
            self.sheets_group.setEnabled(True)
            
            self.statusBar().showMessage(f'{len(self.sheets)} sayfa bulundu')
        except Exception as e:
            QMessageBox.critical(self, 'Hata', f'Sayfalar yüklenirken hata oluştu: {str(e)}')
    
    def select_all_sheets(self):
        """Tüm sayfaları seçer"""
        for i in range(self.sheets_list.count()):
            self.sheets_list.item(i).setSelected(True)
    
    def select_no_sheets(self):
        """Hiçbir sayfayı seçmez"""
        for i in range(self.sheets_list.count()):
            self.sheets_list.item(i).setSelected(False)
    
    def select_output(self):
        """JSON çıktı dosyasını seçme işlemi"""
        # Varsayılan dosya adı
        default_name = ""
        if self.excel_file:
            base_name = os.path.splitext(self.excel_file)[0]
            default_name = f"{base_name}.json"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, 'JSON Dosyasını Kaydet', default_name, 'JSON Dosyaları (*.json);;Tüm Dosyalar (*)')
        
        if file_path:
            # .json uzantısını kontrol et ve ekle
            if not file_path.lower().endswith('.json'):
                file_path += '.json'
            
            self.output_label.setText(os.path.basename(file_path))
    
    def get_selected_sheets(self):
        """Seçili sayfaları döndürür"""
        selected_items = self.sheets_list.selectedItems()
        return [item.text() for item in selected_items]
    
    def convert(self):
        """Excel'i JSON'a dönüştürme işlemi"""
        # Girdi kontrolü
        if not self.excel_file:
            QMessageBox.warning(self, 'Uyarı', 'Lütfen önce bir Excel dosyası seçin.')
            return
        
        selected_sheets = self.get_selected_sheets()
        if not selected_sheets:
            QMessageBox.warning(self, 'Uyarı', 'Lütfen en az bir sayfa seçin.')
            return
        
        # Çıktı dosyası kontrolü
        output_file = self.output_label.text()
        if output_file == 'Henüz seçilmedi':
            # Varsayılan çıktı dosyası adını kullan
            base_name = os.path.splitext(self.excel_file)[0]
            output_file = f"{base_name}.json"
        else:
            # Tam yol oluştur
            output_dir = os.path.dirname(self.excel_file)
            output_file = os.path.join(output_dir, output_file)
        
        # İlerleme çubuğunu göster
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage('Dönüştürülüyor...')
        
        # Dönüştürme işlemini başlat
        self.conversion_thread = ConversionThread(self.excel_file, selected_sheets, output_file)
        self.conversion_thread.progress_signal.connect(self.update_progress)
        self.conversion_thread.finished_signal.connect(self.conversion_finished)
        self.conversion_thread.error_signal.connect(self.conversion_error)
        self.conversion_thread.start()
    
    def update_progress(self, value):
        """İlerleme çubuğunu günceller"""
        self.progress_bar.setValue(value)
    
    def conversion_finished(self, output_file):
        """Dönüştürme işlemi tamamlandığında çağrılır"""
        self.statusBar().showMessage('Dönüştürme tamamlandı!')
        
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle('Başarılı')
        msg_box.setText(f'Excel dosyası başarıyla JSON formatına dönüştürüldü.')
        msg_box.setInformativeText(f'Dosya konumu: {output_file}')
        
        # Dosyayı aç butonu
        open_btn = msg_box.addButton('Dosyayı Aç', QMessageBox.ActionRole)
        close_btn = msg_box.addButton('Kapat', QMessageBox.RejectRole)
        
        msg_box.exec_()
        
        # Kullanıcı "Dosyayı Aç" butonuna tıkladıysa
        if msg_box.clickedButton() == open_btn:
            # Platformdan bağımsız dosya açma
            if sys.platform == 'win32':
                os.startfile(output_file)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{output_file}"')
            else:  # Linux
                os.system(f'xdg-open "{output_file}"')
    
    def conversion_error(self, error_msg):
        """Dönüştürme sırasında hata oluştuğunda çağrılır"""
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage('Hata oluştu')
        
        QMessageBox.critical(self, 'Dönüştürme Hatası', 
                            f'Dönüştürme sırasında bir hata oluştu:\n{error_msg}')

def main():
    app = QApplication(sys.argv)
    # Tema ayarları
    app.setStyle('Fusion')
    window = ExcelToJsonApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()