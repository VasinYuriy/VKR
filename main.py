from copy import copy
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMainWindow
import sys
from UI.loadSyllabusWindow import Ui_loadSyllabusWindow
from excel.examUnitReader import SyllabusReader
from excel.groupReader import GroupReader
from UI.checkForm import Ui_checkForm
from UI.studentChecks import Ui_studentChecks
from UI.examCheckWindow import Ui_examCheckWindow
from UI.examCheckWidget import Ui_examCheck
from UI.global_info import Ui_Global_info
from UI.examPairWindow import Ui_examPairs
from UI.examPairWidget import Ui_examPair
from UI.languageChoice import Ui_languageChoice
from UI.languageChoiceWidget import Ui_languageChoiceWidget
from math import ceil
from word.word import FillTemplate


class Window1(QWidget, Ui_loadSyllabusWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.choose_file)
        self.syllabus_path = self.path_label.text()
        self.exam_units = None
        self.global_units = None
        self.nextButton.clicked.connect(self.read_file)
        self.nextButton.clicked.connect(self.hide)

    def choose_file(self):
        file_name = QFileDialog.getOpenFileName(self, 'Выберите файл...', '', 'Excel file (*.xlsx)')
        self.path_label.setText(file_name[0])
        self.syllabus_path = file_name[0]

    def read_file(self):
        reader = SyllabusReader(self.syllabus_path)
        self.exam_units, self.global_units = reader.read_syllabus()


class Window2(QWidget, Ui_loadSyllabusWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.choose_file)
        self.group_path = None
        self.group_info = None
        self.nextButton.clicked.connect(self.read_file)
        self.nextButton.clicked.connect(self.close)
        self.setWindowTitle('Загрузка оценок по группе')

    def choose_file(self):
        file_name = QFileDialog.getOpenFileNames(self, 'Выберите файл...', '', 'Excel file (*.xlsx)')
        self.group_path = file_name[0]
        path_name = '\n'.join(self.group_path)
        self.path_label.setText(path_name)

    def read_file(self):
        reader = GroupReader(self.group_path)
        self.group_info = reader.get_info()


class StudentChecksWidget(QWidget, Ui_studentChecks):
    def __init__(self, student):
        super().__init__()
        self.setupUi(self)
        self.studentName.setText(student)
        self.document.view().setMinimumWidth(300)


class ExamPairWidget(QWidget, Ui_examPair):
    def __init__(self, exam, exam_list):
        super().__init__()
        self.setupUi(self)
        self.examName.setText(exam)
        self.examChoose.addItems(exam_list)
        self.examChoose.view().setMinimumWidth(600)


class Window3(QWidget, Ui_examPairs):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Несоответствия названий')
        self.nextButton.clicked.connect(self.hide)
        self.widgets = []

    def set_exam_widgets(self, unit_exams, group_exams):
        group_exams_copy = copy(group_exams)
        for exam in group_exams_copy:
            if exam in unit_exams:
                unit_exams.remove(exam)
                group_exams.remove(exam)

        for exam in group_exams:
            self.widgets.append(ExamPairWidget(exam, unit_exams))
            self.verticalLayout.addWidget(self.widgets[-1])

    def get_pairs(self):
        pairs = {}
        for pair in self.widgets:
            pairs[pair.examName.text()] = pair.examChoose.currentText()

        return pairs


def date_converter(month_list):
    months_dict = {
        '01': 'января',
        '02': 'февраля',
        '03': 'марта',
        '04': 'апреля',
        '05': 'мая',
        '06': 'июня',
        '07': 'июля',
        '08': 'августа',
        '09': 'сентября',
        '10': 'октября',
        '11': 'ноября',
        '12': 'декабря',
    }

    return '{} {} {} года'.format(month_list[0], months_dict[month_list[1]], month_list[2])


class Window4(QWidget, Ui_checkForm):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.widgets = []
        self.nextButton.clicked.connect(self.hide)
        self.setWindowTitle('Студенты')
        self.nextButton.clicked.connect(self.get_checks)
        self.students_info = None

    def setup_students(self, students):
        for student in students:
            self.widgets.append(StudentChecksWidget(student))
            self.verticalLayout.addWidget(self.widgets[-1])

    def get_checks(self):
        students_dict = {}

        for widget in self.widgets:
            name = widget.studentName.text()
            student_dict = {
                'Форма обучения': widget.formCheck.isChecked(),
                'Сочетание форм обучения': widget.formSeqCheck.isChecked(),
                'Ускоренное обучение': widget.speedrunCheck.isChecked(),
                'Часть обучения прошла в другой организации': widget.imposterCheck.isChecked(),
                'Факультативные дисциплины': widget.facCheck.isChecked(),
                'Дата рождения': date_converter(widget.birthDate.date().toString("dd MM yyyy").split(' ')),
                'Документ': widget.document.currentText(),
                'Дата выдачи аттестата': widget.certificateYear.text(),
                'Регистрационный номер': widget.studentNumber.text(),
                'Стартап': widget.startupCheck.isChecked()

            }
            students_dict[name] = student_dict
            # print(student_dict['Дата рождения'])
        self.students_info = students_dict


class LanguageChoiceWidget(QWidget, Ui_languageChoiceWidget):
    def __init__(self, student):
        super().__init__()
        self.setupUi(self)
        self.studentName.setText(student)


class Window5(QWidget, Ui_languageChoice):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Выбор языка')
        self.pushButton.clicked.connect(self.hide)
        self.widgets = []

    def set_widgets(self, students):
        for student in students:
            self.widgets.append(LanguageChoiceWidget(student))
            self.widgetsLayout.addWidget(self.widgets[-1])

    def get_languages(self):
        languages = {}
        diploma_languages = {}
        for widget in self.widgets:
            languages[widget.studentName.text()] = widget.languageChoice.currentText()
            diploma_languages[widget.studentName.text()] = widget.languageChoiceDiploma.currentText()
        return languages, diploma_languages


class ExamCheckWidget(QWidget, Ui_examCheck):
    def __init__(self, exam):
        super().__init__()
        self.setupUi(self)
        self.examName.setText(exam)


class Window6(QWidget, Ui_examCheckWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Дисциплины')
        self.widgets = []
        self.nextButton.clicked.connect(self.hide)
        self.nextButton.clicked.connect(self.get_exams)
        self.neededExams = None

    def setup_exams(self, exams):
        for exam in exams:
            self.widgets.append(ExamCheckWidget(exam))
            self.verticalLayout.addWidget(self.widgets[-1])

    def get_exams(self):
        exam_list = [widget.examName.text() for widget in self.widgets if widget.checkBox.isChecked()]
        self.neededExams = exam_list


class Window7(QWidget, Ui_Global_info):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Общая информация')
        self.nextButton.clicked.connect(self.get_global_data)
        self.global_data = None

    def get_global_data(self):

        global_data = {
            'vice_rector': self.viceRectorInfo.toPlainText(),
            'direction': self.directionInfo.toPlainText(),
            'qualification': self.qualificationInfo.currentText(),
            'orientation': self.orientationInfo.toPlainText(),
            'diplomaDate': self.diplomaDate.toPlainText(),
            'protocol': self.protocol.toPlainText(),
        }
        # print(global_data)

        self.global_data = global_data


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.subject_list = None
        self.w1 = Window1()
        self.w2 = Window2()
        self.w3 = Window3()
        self.w4 = Window4()
        self.w5 = Window5()
        self.w6 = Window6()
        self.w7 = Window7()
        self.w1.nextButton.clicked.connect(self.w2.show)
        self.w2.nextButton.clicked.connect(self.w3.show)
        self.w2.nextButton.clicked.connect(self.setup_w3)
        self.w3.nextButton.clicked.connect(self.w4.show)
        self.w3.nextButton.clicked.connect(self.setup_w4)
        self.w4.nextButton.clicked.connect(self.w5.show)
        self.w4.nextButton.clicked.connect(self.add_exams)
        self.w4.nextButton.clicked.connect(self.setup_w5)
        self.w5.pushButton.clicked.connect(self.setup_w6)
        self.w5.pushButton.clicked.connect(self.w6.show)
        self.w6.nextButton.clicked.connect(self.w7.show)
        self.w7.nextButton.clicked.connect(self.w7_next_button_funcs)

    def make_group(self):
        languages, diploma_languages = self.w5.get_languages()
        group = self.Group(self.w4.students_info, self.w1.exam_units, self.w7.global_data, self.w2.group_info, self.w6.neededExams, self.w1.global_units, languages, diploma_languages)
        return group

    class Group:
        def __init__(self, students_info, exam_units, global_data, group_info, needed_exams, global_units, languages, diploma_languages):
            self.students = []
            self.languages = languages
            self.diploma_languages = diploma_languages
            self.exam_units = exam_units
            self.protocol = global_data['protocol']
            self.diplomaDate = global_data['diplomaDate']
            self.orientation = global_data['orientation']
            self.qualification = global_data['qualification']
            self.direction = global_data['direction']
            self.vice_rector = global_data['vice_rector']
            self.code = group_info['code']
            self.faculty = group_info['faculty']
            self.form = group_info['form']
            self.name = group_info['group']
            group_marks = group_info['group_marks']
            course_projects = group_info['course_projects']
            course_works = group_info['course_works']
            excel_dict = group_info['excel_dict']
            self.set_students(
                group_marks=group_marks,
                course_projects=course_projects,
                course_works=course_works,
                students_info=students_info,
                needed_exams=needed_exams,
                exam_units=exam_units,
                global_units=global_units,
                excel_dict=excel_dict
            )

        def set_students(self, group_marks, course_projects, course_works, students_info, needed_exams, exam_units, global_units, excel_dict):
            for student, exams in group_marks.items():
                student_marks = {
                    k: v for k, v in exams.items() if (k.strip() in needed_exams) and (str(v) in [
                        '3', '4', '5', 'зачет', 'None', 'не явился', 'не допущен'])
                }
                print(student)
                student_info = students_info[student]
                language = self.languages[student]
                diploma_language = self.diploma_languages[student]
                self.students.append(self.Student(student, student_marks, course_projects[student], course_works[student], student_info, exam_units, global_units, excel_dict[student], self.qualification, language, diploma_language))

        class Student:
            def __init__(self, name, student_marks, course_projects, course_works, student_info, exam_units, global_units, excel, qualification, language, diploma_language):
                self.secondName = name.split(' ')[0]
                self.firstName = name.split(' ')[1]
                self.thirdName = ' '.join(name.split(' ')[2:])
                self.formCheck = student_info['Форма обучения']
                self.seqFormCheck = student_info['Сочетание форм обучения']
                self.speedrunCheck = student_info['Ускоренное обучение']
                self.imposterCheck = student_info['Часть обучения прошла в другой организации']
                self.facCheck = student_info['Факультативные дисциплины']
                self.birthDate = student_info['Дата рождения']
                self.document = student_info['Документ']
                self.certificateYear = student_info['Дата выдачи аттестата']
                self.studentNumber = student_info['Регистрационный номер']
                self.qualification = qualification
                self.language = language
                self.diploma_language = diploma_language
                self.startup = student_info['Стартап']
                self.first_page_exams, self.second_page_exams = self.format_exams(exam_units, student_marks, course_projects, course_works, global_units)
                self.excel = excel

            def __str__(self):
                return self.secondName

            def format_exams(self, exam_units, student_marks, course_projects, course_works, global_units):
                first_page = []
                second_page = []
                row = 0
                current_list = first_page
                list_flag = False
                mark_dict = {5: 'отлично', 4: 'хорошо', 3: 'удовлетворительно', 'зачет': 'зачтено'}
                max_row = 53
                for exam in exam_units['Дисциплины']:
                    if exam in student_marks.keys():
                        row += ceil(len(exam) / 50)
                        if row >= max_row and not list_flag:
                            row = 0
                            current_list = second_page
                            list_flag = True
                        if student_marks[exam] in mark_dict.keys():
                            mark = mark_dict[student_marks[exam]]
                        else:
                            mark = 'x'
                        # print('{}: {}'.format(exam, row))
                        unit = exam_units['Дисциплины'][exam]
                        if unit != 'x':
                            unit = '{} з. е.'.format(unit)
                        current_list.append([exam, unit, mark])

                row += 4
                if row >= max_row and not list_flag:
                    row = 0
                    current_list = second_page
                    list_flag = True
                else:
                    current_list.append(['', '', ''])
                    row -= 1

                current_list.append(['Практики', '{} з. е.'.format(global_units['Практики']), 'x'])
                current_list.append(['в том числе:', '', ''])
                for exam in exam_units['Практики']:
                    print(student_marks.keys())
                    if exam in student_marks.keys():
                        row += ceil(len(exam) / 50)
                        if row >= max_row and not list_flag:
                            row = 0
                            current_list = second_page
                            list_flag = True

                        if student_marks[exam] in mark_dict.keys():
                            mark = mark_dict[student_marks[exam]]
                        else:
                            mark = 'x'
                        # print('{}: {}'.format(exam, row))
                        unit = exam_units['Практики'][exam]
                        if unit != 'x':
                            unit = '{} з. е.'.format(unit)
                        current_list.append([exam, unit, mark])

                if self.startup:
                    row += 6
                else:
                    row += 5

                if row >= max_row and not list_flag:
                    row = 0
                    current_list = second_page
                    list_flag = True
                else:
                    current_list.append(['', '', ''])
                    row -= 1

                current_list.append([
                    'Государственная итоговая аттестация',
                    '{} з. е.'.format(global_units['Государственная итоговая аттестация']),
                    'x'
                ])
                current_list.append(['в том числе:', '', ''])
                current_list.append(['Государственный экзамен', 'x', ''])
                switch = {
                    'Бакалавр': ' (бакалаврская работа) «...» ',
                    'Магистр': ' (магистерская диссертация) «...» ',
                    'Экономист': ' (дипломная работа) «...» ',
                    'Аспирант': ' (кандидатская диссертация) «...» '
                          }
                text_to_append = ''
                if self.startup and self.diploma_language == 'русский':
                    text_to_append = '(Защита в формате «Стартап как диплом»)'
                elif self.startup and self.diploma_language != 'русский':
                    formatted_language = {
                        'английский': 'английском',
                        'немецкий': 'немецком',
                        'французский': 'французском'}[self.diploma_language]
                    text_to_append = '(Защита в формате «Стартап как диплом» на иностранном ({}) языке)'.format(formatted_language)
                elif not self.startup and self.diploma_language != 'русский':
                    formatted_language = {
                        'английский': 'английском',
                        'немецкий': 'немецком',
                        'французский': 'французском'}[self.diploma_language]
                    text_to_append = '(Защита на иностранном ({}) языке)'.format(formatted_language)

                current_list.append(['Выпускная квалификационная работа' + switch[self.qualification] + text_to_append, 'x', ''])

                row += 5
                if row >= max_row and not list_flag:
                    row = 0
                    current_list = second_page
                    list_flag = True
                else:
                    current_list.append(['', '', ''])

                current_list.append([
                    'Объем образовательной программы',
                    '{} з. е.'.format(global_units['Объем образовательной программы']),
                    'x'
                ])
                current_list.append([
                    'в том числе объем контактной работы обучающихся во взаимодействии с преподавателем в академических часах:',
                    '{} ак. час.'.format(global_units['в том числе объем контактной работы обучающихся во взаимодействии с преподавателем в академических часах:']),
                    'x'
                ])

                row += 2
                if row >= max_row and not list_flag:
                    row = 0
                    current_list = second_page
                    list_flag = True
                else:
                    current_list.append(['', '', ''])
                    row -= 1

                for project, mark in course_projects.items():
                    if mark not in ['не обуч.', 'не выбр.']:
                        row += ceil(len(project) / 50)
                        if row >= max_row and not list_flag:
                            row = 0
                            current_list = second_page
                            list_flag = True

                        if mark in mark_dict.keys():
                            mark = mark_dict[mark]
                        else:
                            mark = 'x'

                        current_list.append(['Курсовой проект, {}'.format(project), 'x', mark])

                for project, mark in course_works.items():
                    if mark not in ['не обуч.', 'не выбр.']:
                        row += ceil(len(project) / 50)
                        if row >= max_row and not list_flag:
                            row = 0
                            current_list = second_page
                            list_flag = True

                        if mark in mark_dict.keys():
                            mark = mark_dict[mark]
                        else:
                            mark = 'x'

                        current_list.append(['Курсовая работа, {}'.format(project), 'x', mark])

                if self.facCheck:
                    row += 4
                    if row >= max_row and not list_flag:
                        row = 0
                        current_list = second_page
                        list_flag = True
                    else:
                        current_list.append(['', '', ''])
                        row -= 1

                    unit_sum = 0
                    for exam in exam_units['Факультативы']:
                        if exam in student_marks.keys():
                            try:
                                unit_sum += exam_units['Факультативы'][exam]
                            except Exception as e:
                                print(e)

                    current_list.append([
                        'Факультативные дисциплины',
                        '',
                        ''
                    ])
                    current_list.append(['в том числе:', '', ''])

                    for exam in exam_units['Факультативы']:
                        if exam in student_marks.keys():
                            row += ceil(len(exam) / 50)
                            if row >= max_row and not list_flag:
                                row = 0
                                current_list = second_page
                                list_flag = True
                            if student_marks[exam] in mark_dict.keys():
                                mark = mark_dict[student_marks[exam]]
                            else:
                                mark = 'x'
                            unit = exam_units['Факультативы'][exam]
                            if unit != 'x':
                                unit = '{} з. е.'.format(unit)
                            current_list.append([exam, unit, mark])

                return first_page, second_page

    def setup_w3(self):
        unit_exams = []
        group_exams = self.w2.group_info['subjects']
        for block in self.w1.exam_units.values():
            unit_exams += block.keys()

        self.w3.set_exam_widgets(unit_exams=unit_exams, group_exams=group_exams)

    def add_exams(self):
        pairs = self.w3.get_pairs()
        for exam1, exam2 in pairs.items():
            for exams in self.w2.group_info['group_marks'].values():
                if exam1 in exams.keys():
                    exams[exam2] = exams[exam1]
                    del exams[exam1]

    def setup_w4(self):
        self.w4.setup_students(self.w2.group_info['group_marks'].keys())

    def setup_w5(self):
        self.w5.set_widgets(self.w2.group_info['group_marks'].keys())

    def setup_w6(self):
        self.w5.get_languages()
        for subject in self.w1.exam_units.values():
            self.w6.setup_exams(subject)

    def make_docs(self):
        filler = FillTemplate(template_path='word/template', group=self.make_group())
        filler.fill_words()

    def w7_next_button_funcs(self):
        self.make_docs()
        app.quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MainWindow()
    window.w1.show()
    app.exec()
