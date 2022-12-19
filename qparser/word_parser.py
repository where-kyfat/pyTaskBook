import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.text.paragraph import Paragraph

from .base_parser import BaseParser


class WordParser(BaseParser):
    @classmethod
    def _parse_question(cls, raw_question: list[Paragraph], chapter: str):
        """ Parse question from list of it's strings

        :param raw_question: list of strings of question and answers in docx.Paragraph type
        :param chapter: name of question chapter
        :return: tuple(question, answers, correct_answer, chapter)
        """

        question = cls._refactor_question(raw_question[0].text)
        answers, correct_answer = cls._refactor_answers(raw_question[1:])

        return question, answers, correct_answer, chapter

    @classmethod
    def _refactor_question(cls, question: str) -> str:
        """Removes start numeration, any symbols like + and so on.
        Also changes first symbol to uppercase

        :param question:
        :return:
        """

        result = re.sub(r'^[+]?\d+\s?[.]', '', question.strip()).strip()  # Removes start numeration and so on
        return result[0].upper() + result[1:]  # Changes first symbol to uppercase

    @classmethod
    def _refactor_answers(cls, raw_answers: list[Paragraph]) -> (list[str], list[int]):
        """Removes numeration of answers and last symbol (like ';', '.' and etc).
        Also finds correct answers

        :param raw_answers: list of raw answers in docx format
        :return: tuple(answers, correct) -> where answers - list of answers
        and correct - list of correct answers
        """

        result = []
        correct = []
        for number, answer in enumerate(raw_answers):
            cur_string = re.sub(r'^\w[).]\s?[.]?', '', answer.text.strip()).strip()  # Remove numeration of answers
            cur_string = re.sub(r'[.;]$', '', cur_string)  # Remove end symbols (like '.', ';')
            result.append(cur_string)

            for symbol in answer.runs:  # find correct
                if symbol.bold:
                    correct.append(number)
                    break
        return result, correct

    @classmethod
    def _document_to_raw_questions(cls, document: Document) -> (list[list[Paragraph]], str):
        """Parses document to list of raw_question, which in contains: question and answers,
        and chapter name

        :param document: opened document in docx.Document format
        :return: tuple(raw_questions, chapter)
        """

        chapter = document.paragraphs[0].text.strip()
        paragraphs = document.paragraphs[2:]  # skip chapter name

        raw_questions = []
        raw_question = []
        for paragraph in paragraphs:
            if paragraph.text.strip():  # if not empty string
                raw_question.append(paragraph)
            else:
                raw_questions.append(raw_question[:])
                raw_question.clear()
        return raw_questions, chapter

    @classmethod
    def parse(cls, path: Path) -> pd.DataFrame:
        """Parses word document of questionnaires to pandas.DataFrame format

        :param path: path to word document
        :return: DataFrame with columns=['Question', 'Answers', 'Correct', 'Chapter']
        """
        document = Document(path)

        raw_questions, chapter = cls._document_to_raw_questions(document)
        questions = [cls._parse_question(raw, chapter) for raw in raw_questions]

        return pd.DataFrame(questions, columns=['Question', 'Answers', 'Correct', 'Chapter'])
