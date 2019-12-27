from django.db import models


class Configuration(models.Model):
    RULES_DIRECTION_TYPES = [("COL", "RULES ARE APPLIED ON THE FIRST COLUMN"),
                             ("ROW", "RULES ARE APPLIED ON THE FIRST ROW")]

    # TODO add date of creation
    name = models.CharField(max_length=50)
    rules_direction = models.CharField(max_length=4, choices=RULES_DIRECTION_TYPES)

    def __str__(self):
        return f'{self.name}'


class Rule(models.Model):
    configuration = models.ForeignKey(Configuration, on_delete=models.CASCADE)
    names = models.TextField(max_length=1000)
    rule = models.CharField(max_length=30)

