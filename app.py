import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame,
    QGridLayout, QPushButton, QHBoxLayout, QLineEdit, QGraphicsDropShadowEffect, QStackedWidget
)
from PyQt5.QtGui import QPixmap, QFont, QIcon, QMovie
from PyQt5.QtCore import Qt, QSize, QTimer
import requests
from io import BytesIO

class InventoryApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Spare-Parts-View V1.0.0")
        self.setGeometry(100, 100, 900, 800)

        # Ler a planilha; títulos começam na linha 4
        self.df = pd.read_excel("../FO-MAN-045 - 03 - LISTA MESTRE - SPARE PARTS.xlsx", engine="openpyxl", skiprows=3)
        self.filtered_df = self.df.copy()
        self.items_per_page = 10
        self.start_index = 0
        self.selected_sector = "GERAL"

        main_layout = QHBoxLayout(self)

        # Menu lateral
        menu_frame = QFrame(self)
        menu_frame.setFixedWidth(200)
        menu_frame.setStyleSheet("background-color: #0e0147;")
        menu_layout = QVBoxLayout(menu_frame)
        menu_layout.setAlignment(Qt.AlignTop)

        # Logo no menu
        logo_label = QLabel(self)
        pixmap = QPixmap("assets/logo.jpg").scaled(120, 120, Qt.KeepAspectRatio)
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignCenter)
        menu_layout.addWidget(logo_label)

        # Botões do menu (armazenados para atualizar o estilo)
        self.menu_buttons = []
        menu_items_list = ["INJEÇÃO", "TRATAMENTO", "MONTAGEM", "PLANTA", "GERAL"]
        for item in menu_items_list:
            btn = QPushButton(item)
            btn.setFixedHeight(40)
            btn.setStyleSheet(self.menu_button_style(default=True))
            btn.clicked.connect(self.filter_by_sector)
            menu_layout.addWidget(btn)
            self.menu_buttons.append(btn)

        # Botão de Sair
        btn_sair = QPushButton("🚪 Sair")
        btn_sair.setFixedHeight(60)
        btn_sair.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 0, 0, 0.8);
                color: white;
                border-radius: 5px;
                font-size: 14px;
                margin-top: 30px;
            }
            QPushButton:hover {
                background-color: rgba(255, 0, 0, 1);
            }
        """)
        btn_sair.clicked.connect(self.close)
        menu_layout.addWidget(btn_sair)

        # Área de conteúdo
        content_frame = QFrame(self)
        content_layout = QVBoxLayout(content_frame)

        # Barra de pesquisa com ícone
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("Pesquisar...")
        self.search_bar.setStyleSheet("""
            QLineEdit {
                background-color: #f1f1f1;
                border-radius: 20px;
                padding: 10px;
                font-size: 14px;
            }
            QLineEdit::placeholder {
                color: #888;
            }
        """)
        self.search_bar.setFixedHeight(40)
        search_icon = QIcon("assets/lupa.png")
        self.search_bar.addAction(search_icon, QLineEdit.LeadingPosition)
        self.search_bar.textChanged.connect(self.filter_items)
        content_layout.addWidget(self.search_bar)

        # QStackedWidget para alternar entre área de rolagem e indicador de carregamento
        self.stack = QStackedWidget()
        
        # Página 0: Conteúdo com os cards (scroll)
        self.scroll = QScrollArea(content_frame)
        self.scroll.setWidgetResizable(True)
        self.scroll_content = QFrame(self.scroll)
        self.grid_layout = QGridLayout(self.scroll_content)
        self.grid_layout.setSpacing(10)
        self.scroll_content.setLayout(self.grid_layout)
        self.scroll.setWidget(self.scroll_content)
        self.stack.addWidget(self.scroll)
        
        # Página 1: Indicador de carregamento (GIF animado)
        self.loading_label = QLabel(self)
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_movie = QMovie("assets/loading.gif")
        self.loading_label.setMovie(self.loading_movie)
        self.loading_label.setVisible(False)
        self.stack.addWidget(self.loading_label)
        
        content_layout.addWidget(self.stack)

        main_layout.addWidget(menu_frame)
        main_layout.addWidget(content_frame)
        self.setLayout(main_layout)

        self.scroll.verticalScrollBar().valueChanged.connect(self.handle_scroll)
        self.load_more_items()

        # Timer para atualizar a planilha periodicamente (ex.: a cada 30 segundos)
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.update_data)
        self.refresh_timer.start(30000)  # 30000 ms = 30 segundos

    def menu_button_style(self, default=True):
        if default:
            return ("""
                QPushButton {
                    background-color: transparent;
                    color: white;
                    border-radius: 5px;
                    font-size: 14px;
                    text-align: left;
                    padding-left: 10px;
                }
                QPushButton:hover {
                    background-color: rgba(255, 255, 255, 0.2);
                }
                QPushButton:pressed {
                    background-color: rgba(255, 255, 255, 0.3);
                }
            """)
        else:
            return ("""
                QPushButton {
                    background-color: #0051a3;
                    color: white;
                    border-radius: 5px;
                    font-size: 14px;
                    text-align: left;
                    padding-left: 10px;
                }
                QPushButton:hover {
                    background-color: #0051a3;
                }
            """)

    def show_loading(self):
        self.stack.setCurrentIndex(1)
        self.loading_movie.setCacheMode(QMovie.CacheAll)
        self.loading_movie.setSpeed(100)
        self.loading_movie.start()
        QApplication.processEvents()

    def hide_loading(self):
        self.loading_movie.stop()
        self.stack.setCurrentIndex(0)

    def handle_scroll(self):
        scrollbar = self.scroll.verticalScrollBar()
        if scrollbar.value() == scrollbar.maximum():
            self.load_more_items()

    def load_more_items(self):
        self.show_loading()  # Exibe o indicador de carregamento
        end_index = self.start_index + self.items_per_page
        data_chunk = self.filtered_df.iloc[self.start_index:end_index]
        self.start_index += self.items_per_page

        row_position = self.grid_layout.rowCount()
        col_position = 0
        for _, row in data_chunk.iterrows():
            codigo = row["CODIGO DA PEÇA"]
            descricao = row["DESCRIÇÃO"]
            quantidade = row["QUANTIDADE"]
            localizacao = row["LOCALIZAÇÃO"]
            img_url = row["IMAGEM"]

            card = QFrame()
            card_layout = QVBoxLayout(card)
            card.setFixedSize(QSize(500, 300))
            card.setStyleSheet("background-color: #ffffff; border-radius: 12px; border: 1px solid #ddd; padding: 10px;")
            
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(10)
            shadow.setXOffset(3)
            shadow.setYOffset(3)
            shadow.setColor(Qt.gray)
            card.setGraphicsEffect(shadow)
            
            img_label = QLabel()
            img_label.setAlignment(Qt.AlignCenter)
            if isinstance(img_url, str) and img_url.startswith("http"):
                try:
                    response = requests.get(img_url)
                    img = QPixmap()
                    img.loadFromData(BytesIO(response.content).read())
                    img_label.setPixmap(img.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                except Exception as e:
                    print(f"Erro ao carregar imagem {img_url}: {e}")
                    default_img = QPixmap("assets/default_image.png")
                    img_label.setPixmap(default_img.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                default_img = QPixmap("assets/default_image.png")
                img_label.setPixmap(default_img.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            
            text_label = QLabel(f"<b>Código:</b> {codigo}<br><b>Descrição:</b> {descricao}<br><b>Quantidade:</b> {quantidade}<br><b>Localização:</b> {localizacao}")
            text_label.setStyleSheet("font-size: 14px; color: #333; padding: 5px;")
            text_label.setFont(QFont("Arial", 10))
            
            card_layout.addWidget(img_label)
            card_layout.addWidget(text_label)
            card.setLayout(card_layout)
            
            self.grid_layout.addWidget(card, row_position, col_position)
            col_position += 1
            if col_position >= 2:
                col_position = 0
                row_position += 1

        # Se nenhum item for exibido, mostra a label "Nenhum item encontrado"
        if self.grid_layout.count() == 0:
            no_label = QLabel("Nenhum item encontrado")
            no_label.setAlignment(Qt.AlignCenter)
            no_label.setStyleSheet("font-size: 16px; color: red; font-weight: bold;")
            self.grid_layout.addWidget(no_label, 0, 0, 1, 2)

        self.hide_loading()

    def filter_by_sector(self):
        self.selected_sector = self.sender().text()
        # Atualiza o estilo dos botões do menu para indicar o setor selecionado
        for btn in self.menu_buttons:
            if btn.text() == self.selected_sector:
                btn.setStyleSheet(self.menu_button_style(default=False))
            else:
                btn.setStyleSheet(self.menu_button_style(default=True))
        if self.selected_sector == "GERAL":
            self.filtered_df = self.df.copy()
        else:
            self.filtered_df = self.df[self.df["SETOR"] == self.selected_sector]
        self.start_index = 0
        self.update_cards()
        self.scroll.verticalScrollBar().setValue(0)

    def update_cards(self):
        # Limpa os cards atuais e recarrega os itens filtrados
        for i in reversed(range(self.grid_layout.count())):
            widget = self.grid_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()
        if self.filtered_df.empty:
            no_label = QLabel("Nenhum item encontrado")
            no_label.setAlignment(Qt.AlignCenter)
            no_label.setStyleSheet("font-size: 16px; color: red; font-weight: bold;")
            self.grid_layout.addWidget(no_label, 0, 0, 1, 2)
        else:
            self.load_more_items()

    def filter_items(self):
        search_text = self.search_bar.text().lower()
        if self.selected_sector != "GERAL":
            df_sector = self.df[self.df["SETOR"] == self.selected_sector]
        else:
            df_sector = self.df.copy()
        if not search_text:
            self.filtered_df = df_sector.copy()
        else:
            self.filtered_df = df_sector[df_sector.apply(lambda row: search_text in str(row["CODIGO DA PEÇA"]).lower() or
                                                          search_text in str(row["DESCRIÇÃO"]).lower(), axis=1)]
        self.start_index = 0
        self.update_cards()
        self.scroll.verticalScrollBar().setValue(0)

    def update_data(self):
        # Atualiza a planilha e, se houver alterações, atualiza os itens
        new_df = pd.read_excel("../FO-MAN-045 - 03 - LISTA MESTRE - SPARE PARTS.xlsx", engine="openpyxl", skiprows=3)
        if not new_df.equals(self.df):
            self.df = new_df.copy()
            if self.selected_sector != "GERAL":
                self.filtered_df = self.df[self.df["SETOR"] == self.selected_sector]
            else:
                self.filtered_df = self.df.copy()
            self.start_index = 0
            self.update_cards()
            self.scroll.verticalScrollBar().setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InventoryApp()
    window.show()
    from PyQt5.QtCore import QTimer
    timer = QTimer(window)
    timer.timeout.connect(window.update_data)
    timer.start(30000)  # Atualiza a cada 30 segundos
    sys.exit(app.exec_())
