from django.db import models
from django.urls import reverse

# Create your models here.

class Product(models.Model):
    code = models.CharField(max_length=50, null=True)
    name = models.CharField(max_length=50)
    quantity = models.IntegerField()
    unit = models.CharField(max_length=50)
    unit_price = models.FloatField()
    total_amount = models.FloatField()

    def get_absolute_url (self):
        return reverse('product_detail', args=[str(self.id)])