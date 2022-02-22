from django.db import models
from django.urls import reverse

# Create your models here.

class Apartment(models.Model):
    so_hop_dong = models.CharField(max_length=50)
    ngay_ky = models.DateField()
    ten_khach_hang = models.CharField(max_length=100)
    email = models.EmailField(max_length=100)
    ngay_hieu_luc = models.DateField()
    ngay_ngay_het_han = models.DateField()
    thoi_han = models.CharField(max_length=50)
    gia_thue = models.IntegerField()
    tien_coc = models.IntegerField()

    def get_absolute_url (self):
        return reverse('apartment_detail', args=[str(self.id)])