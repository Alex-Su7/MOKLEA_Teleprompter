import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QMenuBar, QMenu, QAction, QToolBar, QFileDialog, QVBoxLayout, QWidget, QFontComboBox, QSpinBox, QSlider, QLabel, QHBoxLayout, QMessageBox, QColorDialog, QDialog, QPlainTextEdit
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont, QTextCursor, QColor
import docx

class Teleprompter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('木齐提词器')
        self.setGeometry(100, 100, 800, 600)
        
        # Menu bar
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('文件')
        openAct = QAction('打开', self)
        openAct.triggered.connect(self.openFile)
        fileMenu.addAction(openAct)

        settingsMenu = menubar.addMenu('设置')
        textColorAct = QAction('文本颜色', self)
        textColorAct.triggered.connect(self.changeTextColor)
        settingsMenu.addAction(textColorAct)
        
        bgColorAct = QAction('背景颜色', self)
        bgColorAct.triggered.connect(self.changeBgColor)
        settingsMenu.addAction(bgColorAct)

        helpMenu = menubar.addMenu('帮助')
        instructionsAct = QAction('使用说明', self)
        instructionsAct.triggered.connect(self.showInstructions)
        helpMenu.addAction(instructionsAct)
        
        # Tool bar
        toolbar = self.addToolBar('工具')
        toolbar.addAction(toolbar.toggleViewAction())
        
        # Font family combo box
        self.fontComboBox = QFontComboBox(self)
        self.fontComboBox.currentFontChanged.connect(self.setFontFamily)
        toolbar.addWidget(self.fontComboBox)
        
        # Font size spin box
        self.fontSizeSpinBox = QSpinBox(self)
        self.fontSizeSpinBox.setRange(8, 72)
        self.fontSizeSpinBox.setValue(12)
        self.fontSizeSpinBox.valueChanged.connect(self.setFontSize)
        toolbar.addWidget(self.fontSizeSpinBox)
        
        # Play, pause, and stop buttons
        self.playAct = QAction('播放', self)
        self.playAct.triggered.connect(self.play)
        self.playAct.setShortcut('Ctrl+P')
        toolbar.addAction(self.playAct)
        
        self.pauseAct = QAction('暂停', self)
        self.pauseAct.triggered.connect(self.pause)
        self.pauseAct.setShortcut('Ctrl+S')
        toolbar.addAction(self.pauseAct)
        
        self.stopAct = QAction('终止', self)
        self.stopAct.triggered.connect(self.stop)
        self.stopAct.setShortcut('Ctrl+T')
        toolbar.addAction(self.stopAct)
        
        # Scroll speed slider
        self.speedLabel = QLabel('速度:', self)
        self.speedSlider = QSlider(Qt.Horizontal, self)
        self.speedSlider.setMinimum(1)
        self.speedSlider.setMaximum(100)
        self.speedSlider.setValue(1)
        self.speedSlider.setTickInterval(10)
        self.speedSlider.valueChanged.connect(self.changeSpeed)
        
        self.speedValueLabel = QLabel('1', self)
        
        sliderLayout = QHBoxLayout()
        sliderLayout.addWidget(self.speedLabel)
        sliderLayout.addWidget(self.speedSlider)
        sliderLayout.addWidget(self.speedValueLabel)
        
        sliderWidget = QWidget()
        sliderWidget.setLayout(sliderLayout)
        toolbar.addWidget(sliderWidget)
        
        # Text area
        self.textArea = QTextEdit(self)
        self.textArea.setReadOnly(True)
        
        # Layout
        centralWidget = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.textArea)
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)
        
        # Timer
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.scrollText)
        self.scrollSpeed = 50  # Default scroll speed
        self.scrollIncrement = 1  # Default scroll increment
        
        self.currentFontSize = 12  # Default font size
        self.speedValue = self.speedSlider.value()  # Initialize speed value
        
        self.adjustScrollParameters()  # Initial adjustment of scroll parameters

        self.pausedCursor = None  # Variable to store cursor position when paused

    def openFile(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "打开文本文件", "", "Text Files (*.txt *.docx)", options=options)
        if fileName:
            try:
                if fileName.endswith('.txt'):
                    with open(fileName, 'r') as file:
                        self.textArea.setText(file.read())
                elif fileName.endswith('.docx'):
                    doc = docx.Document(fileName)
                    fullText = [para.text for para in doc.paragraphs]
                    self.textArea.setText('\n'.join(fullText))
            except Exception as e:
                QMessageBox.critical(self, "错误", f"文件读取失败: {e}")

    def play(self):
        if self.pausedCursor:
            self.textArea.setTextCursor(self.pausedCursor)
            self.pausedCursor = None
        self.timer.start(self.scrollSpeed)
        
    def pause(self):
        self.timer.stop()
        self.pausedCursor = self.textArea.textCursor()
        
    def stop(self):
        self.timer.stop()
        self.textArea.moveCursor(QTextCursor.Start)
        self.pausedCursor = None
        
    def scrollText(self):
        cursor = self.textArea.textCursor()
        cursor.movePosition(cursor.Down, cursor.MoveAnchor, self.scrollIncrement)
        self.textArea.setTextCursor(cursor)
    
    def setFontFamily(self, font):
        self.textArea.setFontFamily(font.family())
        self.applyCurrentFont()
        self.adjustScrollParameters()

    def setFontSize(self, size):
        self.currentFontSize = size
        self.applyCurrentFont()
        self.adjustScrollParameters()

    def applyCurrentFont(self):
        cursor = self.textArea.textCursor()
        cursor.select(QTextCursor.Document)
        self.textArea.setTextCursor(cursor)
        self.textArea.setFontPointSize(self.currentFontSize)
        self.textArea.setTextCursor(cursor)
        cursor.clearSelection()

    def changeSpeed(self, value):
        self.speedValue = value
        self.speedValueLabel.setText(str(value))
        self.adjustScrollParameters()
        if self.timer.isActive():
            self.timer.start(self.scrollSpeed)

    def adjustScrollParameters(self):
        # Adjust scroll speed and increment based on font size and slider value
        baseSpeed = 1000  # Base speed, higher values mean slower speed
        speedAdjustment = (self.speedValue - 1) * 8  # Speed adjustment step
        self.scrollSpeed = max(baseSpeed - speedAdjustment, 1)

        # Adjust increment based on font size
        if self.currentFontSize > 24:
            self.scrollIncrement = 1
        elif self.currentFontSize > 18:
            self.scrollIncrement = 2
        elif self.currentFontSize > 12:
            self.scrollIncrement = 3
        else:
            self.scrollIncrement = 4
        
        if self.timer.isActive():
            self.timer.start(self.scrollSpeed)

    def changeTextColor(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.textArea.setTextColor(color)

    def changeBgColor(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.textArea.setStyleSheet(f"background-color: {color.name()};")

    def showInstructions(self):
        instructions = """
        ### 木齐提词器使用说明书

        #### 一、软件介绍
        木齐提词器是一款用于显示提词文本的应用程序，适用于演讲、直播、视频录制等场景。用户可以导入文本文件，并通过该软件自动滚动显示文本，以便于阅读和提示。

        #### 二、安装与运行
        1. 环境要求：
            - 需要安装 Python 3.x。
            - 安装 PyQt5 和 python-docx 库，可以使用以下命令进行安装：
                ```bash
                pip install pyqt5 python-docx
                ```

        2. 运行程序：
            - 将程序代码保存为 `teleprompter.py` 文件。
            - 打开终端或命令提示符，导航到文件所在目录，运行以下命令：
                ```bash
                python teleprompter.py
                ```

        #### 三、界面介绍
        启动程序后，您将看到以下界面元素：

        1. 菜单栏：
            - 文件：包含打开文件选项，允许导入 .txt 和 .docx 文件。
            - 设置：包含文本颜色和背景颜色选项，允许自定义文本和背景的颜色。

        2. 工具栏：
            - 字体选择：下拉菜单，允许选择不同的字体。
            - 字号选择：数字框，允许选择字体大小。
            - 播放：按钮，开始滚动文本。快捷键：Ctrl+P。
            - 暂停：按钮，暂停滚动文本。快捷键：Ctrl+S。
            - 终止：按钮，停止滚动并重置到文本开头。快捷键：Ctrl+T。
            - 速度调节：滑块，调节滚动速度，范围1-100，数字越大速度越快。

        3. 文本区域：
            - 显示导入的文本内容，供用户阅读。

        #### 四、功能说明
        1. 打开文件：
            - 点击菜单栏的“文件”->“打开”，选择要导入的 .txt 或 .docx 文件。
            - 文件内容将显示在文本区域中。

        2. 播放、暂停和终止：
            - 点击工具栏中的“播放”按钮或按 Ctrl+P，开始滚动文本。
            - 点击“暂停”按钮或按 Ctrl+S，暂停滚动。
            - 点击“终止”按钮或按 Ctrl+T，停止滚动并重置到文本开头。

        3. 字体和字号设置：
            - 在工具栏中选择字体下拉菜单，可以更改文本的字体。
            - 在字号选择框中输入或选择字号，可以更改文本的字体大小。

        4. 调整滚动速度：
            - 拖动速度滑块可以调整滚动速度，数字越大滚动速度越快。

        5. 文本颜色和背景颜色设置：
            - 点击菜单栏的“设置”->“文本颜色”，选择颜色对话框中的颜色以更改文本颜色。
            - 点击菜单栏的“设置”->“背景颜色”，选择颜色对话框中的颜色以更改背景颜色。

        #### 五、常见问题与解决
        1. 无法打开文件：
            - 请确保文件格式为 .txt 或 .docx。如果文件较大，可能需要一些时间加载。

        2. 滚动速度调整无效：
            - 请确保在调整速度后程序处于播放状态，滑动速度滑块后，新的速度会立即生效。

        3. 字体和字号设置不生效：
            - 请确保已经选择了文本区域中的字体和字号，若仍有问题，请尝试重新导入文本。

        #### 六、退出程序
        点击窗口右上角的关闭按钮，或者在命令提示符中按 Ctrl+C 结束程序。

        希望本说明书能帮助您更好地使用木齐提词器。如有任何疑问或需要进一步的帮助，请随时联系支持团队。
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("使用说明")
        dialog.setGeometry(100, 100, 600, 400)
        
        layout = QVBoxLayout()
        instructionsText = QPlainTextEdit()
        instructionsText.setPlainText(instructions)
        instructionsText.setReadOnly(True)
        layout.addWidget(instructionsText)
        
        dialog.setLayout(layout)
        dialog.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    teleprompter = Teleprompter()
    teleprompter.show()
    sys.exit(app.exec_())
