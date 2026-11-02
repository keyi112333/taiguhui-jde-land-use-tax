"""读取当前 Windows 运行用户的通用凭据。"""

CREDENTIAL_NOT_FOUND = 1168  # ERROR_NOT_FOUND：指定的通用凭据不存在


class CredentialNotFoundError(RuntimeError):
    """指定的 Windows 通用凭据不存在。"""


def _win_error_code(exc):
    code = getattr(exc, "winerror", None)
    if code is None and getattr(exc, "args", None) and isinstance(exc.args[0], int):
        code = exc.args[0]
    return code


def read_windows_credential(target_name: str):
    if not isinstance(target_name, str) or not target_name.strip():
        raise ValueError("请配置 JDE配置.jde_credential_name（Windows 通用凭据名称）")
    try:
        import win32cred
    except ImportError:
        raise RuntimeError("读取 Windows 凭据需要安装 pywin32") from None
    try:
        credential = win32cred.CredRead(target_name.strip(), win32cred.CRED_TYPE_GENERIC, 0)
    except Exception as exc:
        if _win_error_code(exc) == CREDENTIAL_NOT_FOUND:
            raise CredentialNotFoundError(
                f"Windows 通用凭据不存在：{target_name.strip()}"
            ) from None
        raise RuntimeError("无法读取指定的 Windows 通用凭据，请检查名称及当前运行用户") from None
    username = credential.get("UserName", "")
    password = credential.get("CredentialBlob", b"")
    if isinstance(password, bytes):
        try:
            password = password.decode("utf-16-le").rstrip("\x00")
        except UnicodeDecodeError:
            raise RuntimeError("Windows 凭据密码编码无效，请重新保存通用凭据") from None
    if not username or not password:
        raise RuntimeError("Windows 通用凭据的用户名或密码为空")
    return username, password