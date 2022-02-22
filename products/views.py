from django.urls import reverse_lazy
from django.views.generic import ListView, DetailView
from django.views.generic.edit import CreateView, UpdateView, DeleteView
from .models import Product

class ProductListView(ListView):
    model = Product
    template_name = 'products/product_list.html'
    context_object_name = 'list_all_products' 

    def get_queryset(self):
        code = self.request.GET.get('code', None)
        if code:
            qs = Product.objects.filter(code=code)
            return qs
        return super().get_queryset()

class ProductDetailView(DetailView):
    model = Product
    template_name = 'products/product_detail.html'
    context_object_name = 'product' 

class ProductCreateView(CreateView):
    model = Product
    template_name = 'products/product_create.html'
    fields = ['code', 'name', 'quantity', 'unit', 'unit_price', 'total_amount']
    success_url = reverse_lazy('product_create')

class ProductUpdateView(UpdateView):
    model = Product
    template_name = 'products/product_update.html'
    fields = ['code', 'name', 'quantity', 'unit', 'unit_price', 'total_amount']

class ProductDeleteView(DeleteView):
    model = Product
    template_name = 'products/product_delete.html'
    context_object_name = 'product' 
    success_url = reverse_lazy('product_list')