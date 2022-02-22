from django.urls import reverse_lazy
from django.views.generic import ListView, DetailView
from django.views.generic.edit import CreateView, UpdateView, DeleteView
from .models import Human

class HumanListView(ListView):
    model = Human
    template_name = 'human/human_list.html'
    context_object_name = 'list_all_human' 

    def get_queryset(self):
        cmnd = self.request.GET.get('cmnd', None)
        if cmnd:
            qs = Human.objects.filter(CMND=cmnd)
            return qs
        return super().get_queryset()

class HumanDetailView(DetailView):
    model = Human
    template_name = 'human/human_detail.html'
    context_object_name = 'human' 

class HumanCreateView(CreateView):
    model = Human
    template_name = 'human/human_create.html'
    fields = '__all__'
    success_url = reverse_lazy('human_create')

class HumanUpdateView(UpdateView):
    model = Human
    template_name = 'human/human_update.html'
    fields = '__all__'

class HumanDeleteView(DeleteView):
    model = Human
    template_name = 'human/human_delete.html'
    context_object_name = 'human' 
    success_url = reverse_lazy('human_list')