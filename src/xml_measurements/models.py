from django.db import models


class Configuration(models.Model):
    RULES_DIRECTION_TYPES = [(0, "RULES ARE APPLIED ON THE FIRST COLUMN"), (1, "RULES ARE APPLIED ON THE FIRST ROW")]

    # TODO add date of creation
    name = models.CharField(max_length=50)
    rules_direction = models.IntegerField(choices=RULES_DIRECTION_TYPES)


class Rule(models.Model):
    configuration = models.ForeignKey(Configuration, on_delete=models.CASCADE)
    names = models.TextField(max_length=1000)
    rule = models.CharField(max_length=30)
