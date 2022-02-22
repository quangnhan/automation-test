from django.urls import path
from . import views
 
urlpatterns = [
    path('', views.ApartmentListView.as_view(),name='apartment_list'),
    path('<int:pk>/', views.ApartmentDetailView.as_view(),name='apartment_detail'),
    path('create/', views.ApartmentCreateView.as_view(),name='apartment_create'), 
    path('update/<int:pk>/', views.ApartmentUpdateView.as_view(),name='apartment_update'), 
    path('delete/<int:pk>/', views.ApartmentDeleteView.as_view(),name='apartment_delete'), 
] 