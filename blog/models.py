# models.py
import uuid
from django.db import models


class Users(models.Model):
    email = models.CharField(primary_key=True, max_length=59)
    name = models.CharField(max_length=255)
    password = models.CharField(max_length=42)

class CombinedInfo(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name = models.CharField(max_length=255)
    number = models.CharField(max_length=15, unique=True)
    email = models.EmailField()
    vehicle_number = models.CharField(max_length=255)
    rc_book = models.CharField(max_length=255)
    rc_book_image = models.ImageField(
        upload_to='rc_book_images/', blank=True, null=True)
    previous_policy = models.BooleanField(default=False)
    previous_policy_image = models.ImageField(
        upload_to='previous_policy_images/', blank=True, null=True)
    end_date = models.DateField()

    def _str_(self):
        return f"CombinedInfo for {self.name}"


from django.db import models

class PolicyIssue(models.Model):
    class Meta:
        ordering = ['date', 'name', 'number', 'p_number', 'v_number', 'Vehicle', 'c_number', 'e_number', 'Location', 'HP_bank', 'business_type', 'insurance_type', 'insurance_portal', 'I_company', 'payment', 'payment_sos', 'PS_date', 'PE_date', 'Ncb', 'Premium', 'odNetPremium', 'commissionPercentage', 'payoutDiscount', 'PayoutAmount', 'profitResult', 'tdsPercentage', 'profitAfterTDSResult', 'netProfitResult', 'Executive', 'DSA']
    date = models.DateField(blank=True, null=True)
    name = models.CharField(max_length=255, blank=True, null=True)
    number = models.CharField(max_length=25, blank=True, null=True)
    p_number = models.CharField(max_length=255, blank=True, null=True)
    v_number = models.CharField(max_length=255, blank=True, null=True)
    Vehicle = models.CharField(max_length=255, blank=True, null=True)
    c_number = models.CharField(max_length=255, blank=True, null=True)
    e_number = models.CharField(max_length=255, blank=True, null=True)
    Location = models.CharField(max_length=255, blank=True, null=True)
    HP_bank = models.CharField(max_length=255, blank=True, null=True)
    business_type = models.CharField(max_length=20, choices=[('New', 'New'), ('Data', 'Data'), ('Renewal', 'Renewal'), ('Endorsement', 'Endorsement')], default='Data')
    insurance_type = models.CharField(max_length=20, choices=[('Full', 'Full'), ('TP', 'TP'), ('SOD', 'SOD'), ('Package', 'Package'), ('Health', 'Health')], default='TP')
    insurance_portal = models.CharField(max_length=20, blank=True, null=True, choices=[('Mintpro', 'Mintpro'), ('Agency', 'Agency'), ('Ahmedabad', 'Ahmedabad'), ('IRSS', 'IRSS'), ('Probus', 'Probus'), ('Insurance_Dekho', 'Insurance_Dekho')], default='Agency')
    I_company = models.CharField(max_length=255, blank=True, null=True)
    payment = models.CharField(max_length=255, blank=True, null=True)
    payment_sos = models.CharField(max_length=255, blank=True, null=True)
    PS_date = models.DateField(blank=True, null=True)
    PE_date = models.DateField(blank=True, null=True)
    Ncb = models.IntegerField(blank=True, null=True)
    Premium = models.IntegerField(blank=True, null=True)
    odNetPremium = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    commissionPercentage = models.DecimalField(max_digits=5, decimal_places=2, blank=True, null=True)
    profitResult = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    tdsPercentage = models.DecimalField(max_digits=5, decimal_places=2, blank=True, null=True)
    profitAfterTDSResult = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    payoutDiscount = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    PayoutAmount = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    netProfitResult = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    Executive = models.CharField(max_length=255, null=True)
    DSA = models.CharField(max_length=255, blank=True, null=True)

    def __str__(self):
        return f"{self.name}'s Policy Issue"



class InsuranceEnquiry(models.Model):
    name = models.CharField(max_length=255)
    mobile = models.CharField(max_length=10)
    email = models.EmailField()


class VehicleInformation(models.Model):
    # insurance_enquiry = models.ForeignKey(InsuranceEnquiry, on_delete=models.CASCADE)
    name = models.CharField(max_length=255,null=True)
    mobile = models.CharField(max_length=10,null=True)
    email = models.EmailField(null=True)
    vehicle_number = models.CharField(max_length=255)
    rc_book = models.CharField(max_length=255)
    rc_book_image = models.ImageField(upload_to='rc_book_images/')
    previous_policy = models.CharField(max_length=255)
    previous_policy_image = models.ImageField(
        upload_to='previous_policy_images/')
    end_date = models.DateField()


class LoanEnquiry(models.Model):
    name = models.CharField(max_length=255)
    number = models.CharField(max_length=10)
    email = models.EmailField()


class Document(models.Model):
    loan_enquiry = models.ForeignKey(LoanEnquiry, on_delete=models.CASCADE)
    rc_book = models.CharField(max_length=255)
    rc_book_image = models.ImageField(upload_to='rc_book_images/')
    document = models.FileField(upload_to='documents/')