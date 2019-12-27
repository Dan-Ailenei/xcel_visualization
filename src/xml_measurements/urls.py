from django.urls import path
from . import views
app_name = "xml_measurements"

urlpatterns = [
    path('', views.ConfigurationListView.as_view(), name='index'),
    path('create_configuration/', views.create_update_configuration, name='create_configuration'),
    path('update_configuration/<int:pk>', views.create_update_configuration, name='update_configuration'),
    path('duplicate/<int:pk>', views.duplicate_configuration, name='copy_configuration'),
    path('delete/<int:pk>', views.ConfigurationDeleteView.as_view(), name='delete_configuration'),
    path('inspect/', views.inspect_file_view, name='inspect'),
    path('download/<str:slug>/<int:conf_pk>/<int:sheet_num>', views.download, name='download'),
]
