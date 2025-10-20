# ============================================
# Python 3.13.5
# Software: Spare-Parts-View
# Autor: Rodrigo Almeida
# Data da última atualização: 30/10/2025
# Versão: 1.2.3
#
# ===== NOTAS DE ATUALIZAÇÃO =====
# - Aplicação inicia em geral com botão GERAL selecionado.
# - Aumento de altura dos cards para não cortar as descrições.
# - Mudança de cor do fundo do layout da barra de pesquisa.
# - Funcionalidade de adicionar e visualizar ficha tecnica com otimização de processamento.
# ============================================

import sys, os, io, hashlib, tempfile, pathlib, functools, threading
import pandas as pd, msoffcrypto, requests
from io import BytesIO
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QFrame,
    QGridLayout,
    QPushButton,
    QLineEdit,
    QGraphicsDropShadowEffect,
    QStackedWidget,
    QSizePolicy,
    QMessageBox,
)
from PyQt5.QtGui import QPixmap, QFont, QIcon, QMovie, QDesktopServices
from PyQt5.QtCore import Qt, QSize, QTimer, QUrl

# =========================================================
# CACHE DE IMAGENS
# =========================================================
cache_dir = pathlib.Path(tempfile.gettempdir()) / "spare_parts_cache"
cache_dir.mkdir(exist_ok=True)


@functools.lru_cache(maxsize=512)
def get_pixmap_from_url(url: str) -> QPixmap:
    """Baixa a imagem (timeout 2 s) e salva no cache; retorna QPixmap."""
    h = hashlib.md5(url.encode()).hexdigest() + ".png"
    fp = cache_dir / h
    if fp.exists():
        return QPixmap(str(fp))
    try:
        r = requests.get(url, timeout=2)
        r.raise_for_status()
        pix = QPixmap()
        pix.loadFromData(r.content)
        pix.save(str(fp), "PNG")
        return pix
    except Exception:
        return QPixmap()


# =========================================================
# CLASSE PRINCIPAL
# =========================================================
class InventoryApp(QWidget):
    def __init__(self):
        super().__init__()

         # === Diretório base do app ===
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

        # === Define o ícone da aplicação ===
        icon_path = os.path.join(self.base_dir, "icons", "icone.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"[AVISO] Ícone não encontrado em: {icon_path}")

        # Janela
        self.setWindowTitle("Spare-Parts-View V1.2.3")
        self.setGeometry(100, 100, 1400, 800)
        self.setMinimumSize(1400, 800)

        # Caminhos
        self.base_dir = getattr(
            sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))
        )
        app_dir = (
            os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__))
        )

        self.excel_path = os.path.join(
            app_dir, "FO-MAN-045 - 03 - LISTA MESTRE - SPARE PARTS.xlsm"
        )
        self.senha = "EQUIPEFORTE"

        # Estado
        self.df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        self.items_per_page = 8
        self.start_index = 0
        self.selected_sector = "GERAL"

        # Interface
        self.init_ui()
        self.update_data()

    # =========================================================
    # INTERFACE
    # =========================================================
    def init_ui(self):
        main_layout = QHBoxLayout(self)

        # MENU LATERAL
        menu_frame = QFrame(self)
        menu_frame.setFixedWidth(200)
        menu_frame.setStyleSheet("background-color:#0e0147;")
        menu_layout = QVBoxLayout(menu_frame)
        menu_layout.setAlignment(Qt.AlignTop)

        logo = QLabel(self)
        logo.setPixmap(
            QPixmap(os.path.join(self.base_dir, "icons", "logo.jpg")).scaled(
                120, 120, Qt.KeepAspectRatio
            )
        )
        logo.setAlignment(Qt.AlignCenter)
        menu_layout.addWidget(logo)

        self.menu_buttons = []
        for sector in ["INJEÇÃO", "TRATAMENTO", "MONTAGEM", "PLANTA", "GERAL"]:
            btn = QPushButton(sector)
            btn.setFixedHeight(40)
            btn.setStyleSheet(self.menu_button_style(True))
            btn.clicked.connect(self.filter_by_sector)
            menu_layout.addWidget(btn)
            self.menu_buttons.append(btn)
        self.menu_buttons[-1].setStyleSheet(self.menu_button_style(False))

        sair = QPushButton("\u2794 Sair")
        sair.setFixedHeight(60)
        sair.setStyleSheet(
            "QPushButton{background-color:rgba(255,0,0,.8);color:white;"
            "border-radius:5px;font-size:14px;margin-top:30px}"
            "QPushButton:hover{background-color:rgba(255,0,0,1);}"
        )
        sair.clicked.connect(self.close)
        menu_layout.addWidget(sair)

        rodape = QLabel(
            "Spare Parts View V1.2.3\nDesenvolvido por Rodrigo Almeida\n©NAL BRASIL - 2025"
        )
        rodape.setAlignment(Qt.AlignCenter)
        rodape.setStyleSheet("color:white;font-size:11px;margin-top: 30px;")
        menu_layout.addWidget(rodape)

        # CONTEÚDO PRINCIPAL
        content_frame = QFrame(self)
        content_layout = QVBoxLayout(content_frame)

        # Barra de pesquisa
        search_layout = QHBoxLayout()
        self.search_bar = QLineEdit(self)
        self.search_bar.setFixedHeight(40)
        self.search_bar.setPlaceholderText("Pesquisar...")
        self.search_bar.setStyleSheet(
            "QLineEdit{background:#ffffff;border-radius:20px;padding:10px;font-size:14px}"
            "QLineEdit::placeholder{color:#888;}"
        )
        self.search_bar.addAction(
            QIcon(os.path.join(self.base_dir, "icons", "lupa.png")),
            QLineEdit.LeadingPosition,
        )
        self.search_bar.textChanged.connect(self.filter_items)
        search_layout.addWidget(self.search_bar)

        reload_btn = QPushButton()
        reload_btn.setFixedSize(40, 40)
        reload_btn.setIcon(QIcon(os.path.join(self.base_dir, "icons", "reload.png")))
        reload_btn.setStyleSheet(
            "QPushButton{border:none;}QPushButton:hover{background:rgba(0,0,0,.1);}"
        )
        reload_btn.clicked.connect(self.update_data)
        search_layout.addStretch()
        search_layout.addWidget(reload_btn)
        content_layout.addLayout(search_layout)

        # STACK (lista + loading)
        self.stack = QStackedWidget()
        self.scroll = QScrollArea(content_frame)
        self.scroll.setWidgetResizable(True)
        self.scroll_content = QFrame(self.scroll)
        self.grid_layout = QGridLayout(self.scroll_content)
        self.grid_layout.setSpacing(10)
        self.scroll_content.setLayout(self.grid_layout)
        self.scroll.setWidget(self.scroll_content)
        self.stack.addWidget(self.scroll)

        self.loading_label = QLabel(self)
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_movie = QMovie(
            os.path.join(self.base_dir, "icons", "loading.gif")
        )
        self.loading_label.setMovie(self.loading_movie)
        self.stack.addWidget(self.loading_label)

        content_layout.addWidget(self.stack)
        main_layout.addWidget(menu_frame)
        main_layout.addWidget(content_frame)
        self.setLayout(main_layout)

        # Scroll dinâmico
        self.scroll.verticalScrollBar().valueChanged.connect(self.handle_scroll)

        # Atualização automática
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.update_data)
        self.refresh_timer.start(600_000)

    # =========================================================
    # ESTILOS
    # =========================================================
    def menu_button_style(self, default=True):
        bg = "transparent" if default else "#0051a3"
        hl = "rgba(255,255,255,.2)" if default else "#0051a3"
        return (
            f"QPushButton{{background:{bg};color:white;border-radius:5px;font-size:14px;text-align:left;padding-left:10px}}"
            f"QPushButton:hover{{background:{hl};}}"
        )

    # =========================================================
    # FUNÇÕES DE CARREGAMENTO
    # =========================================================
    def show_loading(self):
        self.stack.setCurrentIndex(1)
        self.loading_movie.start()
        QApplication.processEvents()

    def hide_loading(self):
        self.loading_movie.stop()
        self.stack.setCurrentIndex(0)

    def handle_scroll(self):
        if (
            self.scroll.verticalScrollBar().value()
            == self.scroll.verticalScrollBar().maximum()
        ):
            self.load_more_items()

    # =========================================================
    # CARREGAMENTO DE PLANILHA
    # =========================================================
    def load_data(self):
        try:
            with open(self.excel_path, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                of.load_key(password=self.senha)
                dec = io.BytesIO()
                of.decrypt(dec)
            cols = [
                "ID",
                "CODIGO DA PEÇA\n(Code Number)",
                "DESCRIÇÃO",
                "SETOR",
                "LOCALIZAÇÃO (Kanban Location)",
                "INVENTARIO ATUAL (Actual Inventory)",
                "MÁQUINA (Machine)",
                "FABRICANTE - FORNECEDOR (Manufacturer - Supplier)",
                "PRIORIDADE  (Rank)",
                "IMAGEM",
                "DATASHEET",
            ]
            df = pd.read_excel(dec, engine="openpyxl", skiprows=4, usecols=cols)
            df.columns = df.columns.str.strip()
            df.rename(
                columns={
                    "CODIGO DA PEÇA\n(Code Number)": "CODIGO_DA_PECA",
                    "DESCRIÇÃO": "DESCRIACAO",
                },
                inplace=True,
            )
            self.df = df
        except Exception as e:
            print("Erro ao carregar planilha:", e)
            self.df = pd.DataFrame()

    # =========================================================
    # LISTAGEM DE ITENS
    # =========================================================
    def load_more_items(self):
        self.show_loading()
        end = self.start_index + self.items_per_page
        chunk = self.filtered_df.iloc[self.start_index : end]
        self.start_index = end

        r = self.grid_layout.rowCount()
        c = 0
        for _, row in chunk.iterrows():
            card = QFrame()
            card.setFixedSize(QSize(500, 500))
            card.setStyleSheet(
                "background:white;border:1px solid #ddd;border-radius:12px;padding:10px;"
            )
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(8)
            shadow.setColor(Qt.gray)
            card.setGraphicsEffect(shadow)
            layout = QVBoxLayout(card)

            img_lbl = QLabel()
            img_lbl.setAlignment(Qt.AlignCenter)
            path_or_url = row.get("IMAGEM", "")
            pix = QPixmap()

            if isinstance(path_or_url, str) and path_or_url:
                if path_or_url.lower().startswith("http"):
                    pix = get_pixmap_from_url(path_or_url)
                else:
                    local = (
                        path_or_url
                        if os.path.isabs(path_or_url)
                        else os.path.join(self.base_dir, path_or_url)
                    )
                    if os.path.exists(local):
                        pix = QPixmap(local)
            if pix.isNull():
                pix = QPixmap(
                    os.path.join(self.base_dir, "icons", "default_image.png")
                )
            img_lbl.setPixmap(
                pix.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            )

            txt = QLabel()
            txt.setTextFormat(Qt.RichText)
            txt.setWordWrap(True)
            txt.setFont(QFont("Arial", 10))
            txt.setSizePolicy(
                txt.sizePolicy().horizontalPolicy(), QSizePolicy.Preferred
            )
            txt.setText(
                f"<b>NBA:</b> {row.get('ID','')}<br>"
                f"<b>Código:</b> {row.get('CODIGO_DA_PECA','')}<br>"
                f"<b>Descrição:</b> {row.get('DESCRIACAO','')}<br>"
                f"<b>Quantidade:</b> {row.get('INVENTARIO ATUAL (Actual Inventory)','')}<br>"
                f"<b>Localização:</b> {row.get('LOCALIZAÇÃO (Kanban Location)','')}<br>"
                f"<b>Máquina:</b> {row.get('MÁQUINA (Machine)','')}<br>"
                f"<b>Fabricante:</b> {row.get('FABRICANTE - FORNECEDOR (Manufacturer - Supplier)','')}<br>"
                f"<b>Rank:</b> {row.get('PRIORIDADE  (Rank)','')}"
            )

            layout.addWidget(img_lbl)
            layout.addWidget(txt)

            button_layout = QHBoxLayout()
            button_layout.addStretch()

            action_btn = QPushButton()
            action_btn.setFixedSize(QSize(50, 50))
            action_btn.setIcon(QIcon(os.path.join(self.base_dir, "icons", "sheet.png")))
            action_btn.setIconSize(QSize(40, 40))
            action_btn.setStyleSheet(
                "QPushButton{background-color:#FFFFFF;color:white;border-radius:15px;}"
                "QPushButton:hover{background-color:#AAAAAA;}"
            )
            action_btn.setToolTip("Visualizar ficha técnica")
            action_btn.clicked.connect(
                functools.partial(self.show_details_async, row.get("ID", ""))
            )

            button_layout.addWidget(action_btn)
            layout.addLayout(button_layout)

            self.grid_layout.addWidget(card, r, c)
            c += 1
            if c >= 2:
                c = 0
                r += 1
        self.hide_loading()

    # =========================================================
    # VISUALIZAÇÃO DE FICHA TÉCNICA
    # =========================================================
    def show_details_async(self, item_id):
        """Executa show_details em thread segura, sem travar a UI."""
        def run():
            try:
                self.show_details(item_id)
            finally:
                QTimer.singleShot(0, self.hide_loading)

        item_row = self.df[self.df["ID"] == item_id]
        if item_row.empty:
            self.show_message("Item não encontrado.")
            return

        datasheet_value = item_row["DATASHEET"].iloc[0]
        datasheet_path = str(datasheet_value).strip() if isinstance(datasheet_value, (str, bytes)) else ""

        # Evita abrir thread se não houver ficha técnica
        if not datasheet_path:
            self.show_message("Este item não tem ficha técnica cadastrada.")
            return

        # Exibe loading e inicia thread
        self.show_loading()
        thread = threading.Thread(target=run)
        thread.start()

    def show_details(self, item_id):
        """Abre o link da ficha técnica (web, local, rede ou relativa à pasta datasheet)."""
        try:
            item_row = self.df[self.df["ID"] == item_id]
            if item_row.empty:
                return

            datasheet_value = item_row["DATASHEET"].iloc[0]
            datasheet_path = str(datasheet_value).strip() if isinstance(datasheet_value, (str, bytes)) else ""

            if not datasheet_path:
                return

            if datasheet_path.startswith(("http://", "https://")):
                QDesktopServices.openUrl(QUrl(datasheet_path))
                return

            if datasheet_path.startswith(("\\\\", "file://")):
                QDesktopServices.openUrl(QUrl.fromUserInput(datasheet_path))
                return

            if os.path.isabs(datasheet_path) and os.path.exists(datasheet_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(datasheet_path))
                return

            possible_path = os.path.abspath(os.path.join(self.base_dir, "..", "datasheet", datasheet_path))
            if os.path.exists(possible_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(possible_path))
                return

            self.show_message(f"O arquivo da ficha técnica não foi encontrado:\n\n{datasheet_path}")

        except Exception as e:
            self.show_message(f"Erro ao abrir ficha técnica:\n{e}")

    # =========================================================
    # FILTROS E ATUALIZAÇÃO
    # =========================================================
    def show_message(self, text):
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Ficha Técnica")
        msg_box.setText(text)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def filter_by_sector(self):
        self.selected_sector = self.sender().text()
        for b in self.menu_buttons:
            b.setStyleSheet(self.menu_button_style(b.text() != self.selected_sector))
        self.filtered_df = (
            self.df
            if self.selected_sector == "GERAL"
            else self.df[self.df["SETOR"] == self.selected_sector]
        )
        self.start_index = 0
        self.update_cards()

    def filter_items(self):
        txt = self.search_bar.text().lower()
        df_sector = (
            self.df
            if self.selected_sector == "GERAL"
            else self.df[self.df["SETOR"] == self.selected_sector]
        )
        if txt:
            cols = [
                "CODIGO_DA_PECA",
                "DESCRIACAO",
                "ID",
                "MÁQUINA (Machine)",
                "FABRICANTE - FORNECEDOR (Manufacturer - Supplier)",
            ]
            mask = pd.concat(
                [
                    df_sector[c].astype(str).str.lower().str.contains(txt, na=False)
                    for c in cols
                ],
                axis=1,
            ).any(axis=1)
            self.filtered_df = df_sector[mask]
        else:
            self.filtered_df = df_sector.copy()
        self.start_index = 0
        self.update_cards()

    def update_data(self):
        self.load_data()
        self.filtered_df = (
            self.df
            if self.selected_sector == "GERAL"
            else self.df[self.df["SETOR"] == self.selected_sector]
        )
        self.start_index = 0
        self.update_cards()

    def update_cards(self):
        for i in reversed(range(self.grid_layout.count())):
            w = self.grid_layout.itemAt(i).widget()
            if w:
                w.deleteLater()
        if self.filtered_df.empty:
            lbl = QLabel("Nenhum item encontrado")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("color:red;font-size:20px;font-weight:bold;")
            self.grid_layout.addWidget(lbl, 0, 0, 1, 2)
        else:
            self.load_more_items()
        self.scroll.verticalScrollBar().setValue(0)

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = InventoryApp()
    win.show()
    sys.exit(app.exec_())