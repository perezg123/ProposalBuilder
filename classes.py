import datetime
from dataclasses import dataclass

@dataclass
class Product:
    id: str
    unit: str
    sku: str
    description: str
    price: float
    one: float
    two: float
    three: float
    four: float
    five: float
    comments: str
    category: str