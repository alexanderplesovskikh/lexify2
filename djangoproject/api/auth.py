from api.models import User
from rest_framework.exceptions import AuthenticationFailed

def authenticate_token(request):
    token = request.headers.get('Authorization')
    if not token:
        raise AuthenticationFailed("Token required")
    try:
        return User.objects.get(token=token)
    except User.DoesNotExist:
        raise AuthenticationFailed("Invalid token")
