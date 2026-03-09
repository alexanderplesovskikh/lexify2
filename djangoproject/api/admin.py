from django.contrib import admin
from django.apps import apps
from django.contrib.auth.models import Group

# Unregister default Group if not needed
# admin.site.unregister(Group)

# Get all models from your app
app_models = apps.get_app_config('api').get_models()

SKIP_MODELS = ['Style', 'Format', 'Dictionary', 'Workspace']

# Custom admin for Docx model
class DocumentAdmin(admin.ModelAdmin):
    list_display = ('id', 'updated_at', 'user__email', 'processed')
    list_filter = ('processed', 'updated_at', 'user__email')
    search_fields = ('id',)
    readonly_fields = ('id', 'updated_at')
    
try:
    from api.models import Document
    admin.site.unregister(Document)
except (ImportError, admin.sites.NotRegistered):
    pass

# Auto-register all models
for model in app_models:
    if model.__name__ not in SKIP_MODELS:
        try:
            # Special handling for Docx model
            if model.__name__ == 'Document':
                print(f"Registering Document with custom admin")  # This will show in console
                admin.site.register(model, DocumentAdmin)
            else:
                # Default registration for all other models
                admin.site.register(model)
        except admin.sites.AlreadyRegistered:
            print(f"Model {model.__name__} already registered")
            pass
        except Exception as e:
            print(f"Error registering {model.__name__}: {e}")