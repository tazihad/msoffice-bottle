### MS Office 2010
Use Microsoft Office 2010 in Linux using WINE and Bottles.  

![Screenshot_20240701_123107](https://github.com/tazihad/msoffice-bottle/assets/19417232/6e51ebef-4e25-4a78-af1e-725c197fc8c2)  

#### Video Tutorial:  

[![YOUTUBE_TUTORIAL](https://img.youtube.com/vi/Gb95naAjBsg/0.jpg)](https://www.youtube.com/watch?v=Gb95naAjBsg)

*Prerequisite:*   
[Flatpak](https://flatpak.org/setup/)  
[Bottles installed by Flatpak](https://flathub.org/apps/com.usebottles.bottles)  

Permissions:  
Give drive permissions using flatseal or do from commandline
```
flatpak override com.usebottles.bottles --user --filesystem=xdg-data/applications
flatpak override com.usebottles.bottles --user --filesystem=home
```

1. Download WINE Runner:
```
mkdir -p ~/.var/app/com.usebottles.bottles/data/bottles/runners/pol-8.2 && \
wget https://www.playonlinux.com/wine/binaries/phoenicis/upstream-linux-x86/PlayOnLinux-wine-8.2-upstream-linux-x86.tar.gz \
-O /tmp/PlayOnLinux-wine-8.2-upstream-linux-x86.tar.gz && \
tar -xz -C ~/.var/app/com.usebottles.bottles/data/bottles/runners/pol-8.2 --strip-components=1 \
-f /tmp/PlayOnLinux-wine-8.2-upstream-linux-x86.tar.gz && \
rm /tmp/PlayOnLinux-wine-8.2-upstream-linux-x86.tar.gz
```

2. Create a new bottle using Bottles named `office2010`  
3. Environment `Custom`.  
4. From Custom `Runner: pol-8.2`. Architecture: `32-bit`.  
5. Configaration: Select `office2010.yml`.  
6. Click `Create`.
   
![Screenshot_20240701_113937](https://github.com/tazihad/msoffice-bottle/assets/19417232/916c186a-08c8-4b81-8504-21ae0bab7dd3)  

> Unfortunately for some bug, DLL override doesn't get imported from yml. So, we do this manually.
7. From bottle `office2010` -> `Setting` -> `DLL Overrides` -> add `gdiplus` & `riched20`.

![Screenshot_20240701_114718](https://github.com/tazihad/msoffice-bottle/assets/19417232/ef064bed-62aa-4349-9424-7188cb4f6cb0)  

8. That's it. Now install 32-bit version of *MS Office 2010*. Find the executable and run with Bottles.  

### File Integrations  
Copy files `word.desktop`, `powerpoint.desktop`, `excel.desktop` files from [office2010-file-integrations](https://github.com/tazihad/msoffice-bottle/tree/main/office2010-file-integrations)  
 to `~/.local/share/applications` Change `Icon` path by changing to your username.  
Or do run this command:  
```
mkdir -p ~/.local/share/applications && \
curl -o ~/.local/share/applications/Office-Excel-2010.desktop https://raw.githubusercontent.com/tazihad/msoffice-bottle/main/office2010-file-integrations/Office-Excel-2010.desktop && \
curl -o ~/.local/share/applications/Office-Powerpoint-2010.desktop https://raw.githubusercontent.com/tazihad/msoffice-bottle/main/office2010-file-integrations/Office-Powerpoint-2010.desktop && \
curl -o ~/.local/share/applications/Office-Word-2010.desktop https://raw.githubusercontent.com/tazihad/msoffice-bottle/main/office2010-file-integrations/Office-Word-2010.desktop
```


### Office 2013, 2016

Wine runner pol-4.3:
```
mkdir -p ~/.var/app/com.usebottles.bottles/data/bottles/runners/pol-4.3 && \
wget https://www.playonlinux.com/wine/binaries/phoenicis/upstream-linux-x86/PlayOnLinux-wine-4.3-upstream-linux-x86.tar.gz \
-O /tmp/PlayOnLinux-wine-4.3-upstream-linux-x86.tar.gz && \
tar -xz -C ~/.var/app/com.usebottles.bottles/data/bottles/runners/pol-4.3 --strip-components=1 \
-f /tmp/PlayOnLinux-wine-4.3-upstream-linux-x86.tar.gz && \
rm /tmp/PlayOnLinux-wine-4.3-upstream-linux-x86.tar.gz
```
