from django.urls import path
from . import views
 
urlpatterns = [
    path('', views.HumanListView.as_view(),name='human_list'),
    path('<int:pk>/', views.HumanDetailView.as_view(),name='human_detail'),
    path('create/', views.HumanCreateView.as_view(),name='human_create'), 
    path('update/<int:pk>/', views.HumanUpdateView.as_view(),name='human_update'), 
    path('delete/<int:pk>/', views.HumanDeleteView.as_view(),name='human_delete'), 
] 
