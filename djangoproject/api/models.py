import uuid
from django.db import models
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager

# ---------- USER ----------

import uuid
from django.db import models
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin

class UserManager(BaseUserManager):
    def create_user(self, email, is_admin=False):
        token = uuid.uuid4().hex
        user = self.model(email=email, token=token, is_admin=is_admin)
        user.set_unusable_password()
        user.save()
        return user
    
    def create_superuser(self, email, password=None, **extra_fields):
        # Create the user
        token = uuid.uuid4().hex
        user = self.model(
            email=email, 
            token=token,
            is_admin=True,
            is_staff=True,
            is_superuser=True,
            **extra_fields
        )
        
        if password:
            user.set_password(password)
        else:
            user.set_unusable_password()
            
        user.save()
        return user


class User(AbstractBaseUser, PermissionsMixin):  # Make sure PermissionsMixin is here!
    email = models.EmailField(unique=True)
    token = models.CharField(max_length=64, unique=True)
    is_admin = models.BooleanField(default=False)
    
    # REQUIRED for Django admin
    is_staff = models.BooleanField(default=False)
    is_active = models.BooleanField(default=True)
    
    # is_superuser comes from PermissionsMixin automatically
    
    USERNAME_FIELD = 'email'
    objects = UserManager()

    def __str__(self):
        return self.email
    
    # Required for admin access
    def has_perm(self, perm, obj=None):
        return self.is_staff or self.is_admin
    
    def has_module_perms(self, app_label):
        return self.is_staff or self.is_admin



# ---------- FILE ----------

class Document(models.Model):

    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    user = models.ForeignKey(User, on_delete=models.CASCADE)

    style = models.CharField(max_length=50, blank=True)
    format = models.CharField(max_length=50, blank=True)
    dictionary = models.CharField(max_length=50, blank=True)
    skip_pages = models.CharField(max_length=50, blank=True)

    name = models.CharField(max_length=1000, blank=True, null=True, default="Document.docx")

    plag = models.TextField(blank=True, null=True, default=" ")

    processed = models.BooleanField(default=False)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    count_chars = models.BigIntegerField(null=True, blank=True, default=0)
    count_words = models.BigIntegerField(null=True, blank=True, default=0)
    count_sentences = models.BigIntegerField(null=True, blank=True, default=0)

    count_bad_words = models.BigIntegerField(null=True, blank=True, default=0)
    count_bad_chars = models.BigIntegerField(null=True, blank=True, default=0)
    count_bibliography = models.BigIntegerField(null=True, blank=True, default=0)
    count_bad_bibliography = models.BigIntegerField(null=True, blank=True, default=0)
    count_not_doi = models.BigIntegerField(null=True, blank=True, default=0)
    count_suggest_doi = models.BigIntegerField(null=True, blank=True, default=0)
    count_not_right_bibliography = models.BigIntegerField(null=True, blank=True, default=0)
    count_styles_error = models.BigIntegerField(null=True, blank=True, default=0)
    

class UserToken(models.Model):
    email = models.EmailField(unique=True, db_index=True, verbose_name="Email")
    token = models.CharField(max_length=64, unique=True, blank=True, null=True, verbose_name="Токен")
    code = models.CharField(max_length=6, blank=True, null=True, verbose_name="Код")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата")

    class Meta:
        verbose_name = "Пользователя"
        verbose_name_plural = "Пользователи"
        db_table = 'docx_usertoken'

    def __str__(self):
        return f'{self.email}'
