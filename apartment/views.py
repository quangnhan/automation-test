from django.urls import reverse_lazy
from django.views.generic import ListView, DetailView
from django.views.generic.edit import CreateView, UpdateView, DeleteView
from .models import Apartment

class ApartmentListView(ListView):
    model = Apartment
    template_name = 'apartment/apartment_list.html'
    context_object_name = 'list_all_apartment' 

    def get_queryset(self):
        so_hop_dong = self.request.GET.get('so_hop_dong', None)
        if so_hop_dong:
            qs = Apartment.objects.filter(so_hop_dong=so_hop_dong)
            return qs
        return super().get_queryset()

class ApartmentDetailView(DetailView):
    model = Apartment
    template_name = 'apartment/apartment_detail.html'
    context_object_name = 'apartment' 


class ApartmentCreateView(CreateView):
    model = Apartment
    template_name = 'apartment/apartment_create.html'
    fields = '__all__'
    success_url = reverse_lazy('apartment_create')

class ApartmentUpdateView(UpdateView):
    model = Apartment
    template_name = 'apartment/apartment_update.html'
    fields = '__all__'

class ApartmentDeleteView(DeleteView):
    model = Apartment
    template_name = 'apartment/apartment_delete.html'
    context_object_name = 'apartment' 
    success_url = reverse_lazy('apartment_list')