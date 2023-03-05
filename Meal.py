# Meal class
class Meal:

    def __init__(self, ingredients):
        self.ingredients = ingredients
    
    def print_ingredients(self):
        for item in self.ingredients:
            print(item)