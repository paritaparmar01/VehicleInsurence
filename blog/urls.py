# blog/urls.py
from django.urls import path
from . import views
from django.conf import settings
from django.conf.urls.static import static
from blog.views import upload_excel

# app_name = 'insurance'

urlpatterns = [
    # path('enquiry/', InsuranceEnquiryView.as_view(), name='enquiry'),
    #path('', views.index, name="base"),
    path('register/', views.register, name="register"),
    path('login/', views.login, name="login"),
    path('forgotpwd/', views.forgotPwd, name="forgotpwd"),
    path('policyissue/', views.policy_issue, name="policy_issue"),
    path('uploadcsvpolisyissue/', views.upload_policy, name="upload_policy"),
    path('uploadexcel/', upload_excel, name='upload_excel'),
    #path('enquiry/', views.enquiry, name="enquiry"),
    #path('enquiryoldcust/', views.enquiryOldCust, name="enquiryoldcust"),
    path('enquirynewcust/', views.enquiryNewCust, name="enquirynewcust"),
    path('loan/', views.loan, name="loan"),
    path('exporttoexcel/', views.export_to_excel, name='export_to_excel'),
    path('', views.home, name='home'),
    path('try/', views.try_button, name='try_button'),
    path('try/search/', views.search, name='search'),
    path('exporttoexcel/table/<Param1>/<Param2>/', views.downloadsingletable, name='downloadsingletable'),
    #path('downloadtable/', views.downloadsingletable, name='downloadtabledata'),
    path('trypolicy/',views.try_policy, name='trypolicy'),
    path('tryloan/',views.try_loan, name='tryloan'),

    
    # path('exporttoexcelzip/', views.export_to_excel_zip, name='export_to_excel_zip')

    # Add more URL patterns as needed
]
# Include static files during development
if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
