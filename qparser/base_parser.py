"""Base parser realization"""
from abc import ABCMeta, abstractmethod

import pandas as pd


class BaseParser(metaclass=ABCMeta):
    """
    Abstract class BaseParser

    Encapsulation read/write logic for future Parser classes
    """

    def __init__(self, path_for_parse: list[str],
                 path_to_save: str,
                 add_to_quest: str = '',
                 save_xlsx: bool = False):
        """Constructor of abc BaseParser

        :param path_for_parse: list of path's to parse
        :param path_to_save: path to save resulting questionnaires
        :param add_to_quest: path to existing questionnaires file for add new questions
        """
        self.path_for_parse = path_for_parse
        self.path_to_save = path_to_save
        self.add_to_quest = add_to_quest
        self.save_xlsx = save_xlsx

    def __call__(self) -> list[str] | str:
        """ Call method of abs BaseParser

        Parses raw questionnaires and save result in file,
        according to configuration
        """
        parse_results = []
        for path in self.path_for_parse:
            parse_results.append(self.parse(path))
        quest = self.merge_quest(parse_results)

        result_paths = []
        if self.save_xlsx:
            result_paths.append(self.write_to_xlsx(quest))

        result_paths.append(self.write_to_json(quest))
        result_paths.append(self.write_to_csv(quest))
        return result_paths

    def write_to_csv(self, quest: pd.DataFrame) -> str:
        """Write quest to resulting csv file

        :param quest: Questionnaire
        :return: path to resulting file
        """
        if self.add_to_quest:
            df_old = pd.read_csv(self.add_to_quest)
            merged = df_old.merge(quest)
            merged.to_csv(self.add_to_quest)
            return self.add_to_quest
        else:
            quest.to_csv(self.path_to_save)
            return str(self.path_to_save)

    def write_to_xlsx(self, quest: pd.DataFrame) -> str:
        """Write quest to xlsx file

        :param quest: Questionnaire
        :return: path to resulting file
        """

        path_to_excel = self.path_to_save[:-3] + 'xlsx'  # removes csv and add xlsx to path
        quest.to_excel(path_to_excel)
        return path_to_excel

    def write_to_json(self, quest: pd.DataFrame) -> str:
        """Write quest to json file

        :param quest: Questionnaire
        :return: path to resulting file
        """

        path_to_json = self.path_to_save[:-3] + 'json'  # removes csv and json to path
        quest.to_json(path_to_json)
        return path_to_json

    @classmethod
    def merge_quest(cls, quests: list[pd.DataFrame]) -> pd.DataFrame:
        """ Merging list of received questionnaires

        :param quests: list of Questionnaires
        :return: merged questionnaire
        """
        df = pd.concat(quests)
        df.reset_index(inplace=True)
        return df

    @abstractmethod
    def parse(self, path: str) -> pd.DataFrame:
        """Parse questionnaire from file

        :param path: path to raw questionnaire file
        """
        pass
