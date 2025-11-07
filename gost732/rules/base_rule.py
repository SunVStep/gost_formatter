from docx import Document
from abc import ABC, abstractmethod

# Определяем базовый класс для всех последующий файлов
class BaseRule(ABC):
    name: str = "Base rule"

# Используем декоратор для обозначения абстрактного класса
    @abstractmethod
    def apply(self, doc:Document) -> None: pass # Выбранное правило будет применяться к документу