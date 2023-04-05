"""

Windows Volume Controller
--------------------------

For controlling master volume using pycaw lib and core win audio endpoints

Get Started -
1. Make sure `chocolatey` and `python` are installed
2. `pip install pycaw`
3. `choco install visualcpp-build-tools`

"""

from ctypes import POINTER, cast

from comtypes import CLSCTX_ALL

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class MasterVolumeController:
    def __init__(self):
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

        # Normalization value (for 0 to 100)
        self.norm = 100

        # Windows Volume Endpoint (https://learn.microsoft.com/en-us/windows/win32/api/endpointvolume/nn-endpointvolume-iaudioendpointvolume)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))

    def is_mute(self) -> bool:
        return self.volume.GetMute() == 1

    def mute(self):
        self.volume.SetMute(1, None)

    def unmute(self):
        self.volume.SetMute(0, None)

    def get_volume_db(self) -> float:
        return self.volume.GetMasterVolumeLevel()

    def get_volume(self) -> float:
        return self.volume.GetMasterVolumeLevelScalar()

    def set_volume_db(self, db: float):
        self.volume.SetMasterVolumeLevel(db, None)

    def set_volume(self, vol: float):
        self.volume.SetMasterVolumeLevelScalar(vol, None)

    def increase_volume(self, steps: int):
        vol = min(self.get_volume() + (steps / self.norm), 1.0)
        self.set_volume(vol=vol)

    def decrease_volume(self, steps: int):
        vol = max(self.get_volume() - (steps / self.norm), 0.0)
        self.set_volume(vol=vol)


def main():
    controller = MasterVolumeController()

    # Get Current Master Volume (0.0 to 1.0)
    print(f"Current Volume: {controller.get_volume()}")

    # Set Master Volume to 50%
    controller.set_volume(0.5)

    # Increment current volume by 25 steps
    controller.increase_volume(40)


if __name__ == "__main__":
    main()
