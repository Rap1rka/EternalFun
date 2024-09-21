from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from comtypes import CLSCTX_ALL

# Функция для восстановления громкости всех аудиосессий на 100%
def unmute_all_sessions():
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        # Устанавливаем громкость на 100% для каждой сессии
        volume.SetMasterVolume(1.0, None)  # 1.0 соответствует 100% громкости

if __name__ == "__main__":
    print("Восстанавливаем громкость для всех сессий...")
    unmute_all_sessions()
    print("Звуки восстановлены до 100% громкости.")
