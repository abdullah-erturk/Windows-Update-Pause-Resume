<a href="https://buymeacoffee.com/abdullaherturk" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" ></a>

# Windows Update Pause / Resume v2 (safe way)

![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011%20%7C%20Server%202012R2%2B-0078D6?logo=windows&logoColor=white)
![Script](https://img.shields.io/badge/Script-PowerShell%20%2B%20Batch-5391FE?logo=powershell&logoColor=white)
![UI](https://img.shields.io/badge/UI-T%C3%BCrk%C3%A7e%20%7C%20English-2ea44f)
![Auto Elevation](https://img.shields.io/badge/Auto-Elevation%20(UAC)-orange)
![Defender](https://img.shields.io/badge/Defender-stays%20updated-44cc11?logo=windows-defender&logoColor=white)

![sample](https://github.com/abdullah-erturk/Windows-Update-Pause-Resume/blob/main/prewiev.gif)

---

## 🇹🇷 Türkçe Açıklama

<details>
<summary><b>🆕 Yenilikler</b></summary>

- ⚙️ **VBS yerine PowerShell** — kod tamamen `VBScript`'ten `PowerShell`'e taşındı. Ancak özel bir **hibrit (polyglot) yöntemle**, PowerShell kodları `.bat` uzantılı tek bir dosya içinde çalışır: dosyanın başındaki batch bölümü, kendi içeriğini okuyup PowerShell'e aktarır (`Invoke-Expression`). Böylece `.ps1` yürütme ilkesiyle uğraşmadan çift tıklamayla çalışır.
- ✅ **Windows 10 / 11 (26H1 dahil) desteği** — yerleşik *Pause Updates* API'si ile, `Pause*StartTime` / `Pause*ExpiryTime` değerleri **gerçek bitiş tarihiyle (UTC)** yazılarak güncelleştirmeler gerçek anlamda duraklatılır.
- 🖥️ **Windows Server 2016 ve eski sürüm desteği** — bu sürümlerde Pause API'si yoktur; betik bunu otomatik algılar ve otomatik OS güncelleştirmelerini **`NoAutoUpdate` ilkesiyle** durdurur. **`wuauserv` hizmeti çalışmaya devam eder.** Belirttiğiniz hafta dolunca ilkeyi geri açan bir **zamanlanmış görev** kurulur.
- 🛡️ **Güvenlik korunur** — her iki yöntemde de yalnızca işletim sistemi güncelleştirmeleri durur; **Microsoft Defender virüs tanımı (imza) güncellemelerini almaya devam eder**, bilgisayarınız korumasız kalmaz.
- 🌐 **Otomatik Türkçe / İngilizce arayüz** — işletim sistemi dili Türkçe ise arayüz Türkçe, değilse İngilizce olur.
- 📦 **Tek dosya** — ayrı TR/ENG sürümlerine gerek yok; aynı betik her iki dili de destekler.
- 🪟 **Konsol başlığı** — açılan pencerenin başlığında `Windows Update Pause / Resume | by Abdullah ERTÜRK` yazar.

</details>

<details>
<summary><b>✨ Özellikler</b></summary>

- **Kolay Kullanım:** Sadece bir kaç çift tıklama ile Windows güncelleştirmelerini duraklatabilirsiniz.
- **Hızlı ve Etkili:** Güncelleştirmeleri anında duraklatır, zaman kaybı yaşatmaz.
- **Esnek Süre:** Duraklatma süresini hafta olarak siz belirlersiniz (**1 - 10000 hafta arası**); geçersiz veya aralık dışı girişler engellenir.
- **Geri Alınabilir:** Betiği tekrar çalıştırarak güncelleştirmeleri kaldığınız yerden devam ettirebilirsiniz.
- **Güvenli:** Windows'un yerleşik komut ve ilkelerini kullanarak güvenli bir şekilde işlem yapar.
- **Defender Korumalı:** Güncelleştirmeler duraklatılsa bile Windows Defender, virüs tanımı (imza) veritabanını güncellemeye devam eder.
- **Geniş Uyumluluk:** Windows 10, Windows 11 ve Windows Server 2016 (ve daha eski sürümler) desteklenir; betik sürümü algılayıp en uygun yöntemi kendisi seçer.
- **Otomatik Dil:** Betik, işletim sisteminizin dilini algılar ve arayüzü Türkçe ya da İngilizce gösterir.

</details>

<details>
<summary><b>⚙️ Nasıl Çalışır</b></summary>

- **Modern Windows 10 / 11 (derleme 15063 ve üzeri):** Windows'un yerleşik "Güncelleştirmeleri Duraklat" özelliği kullanılır; süre dolduğunda güncelleştirmeler kendiliğinden devam eder.
- **Windows Server 2016 ve eski sürümler (derleme 15063 altı):** Pause API'si bulunmadığından otomatik güncelleştirmeler `NoAutoUpdate` ilkesiyle durdurulur, `wuauserv` hizmeti çalışmaya devam eder ve seçtiğiniz süre sonunda ilkeyi geri açan zamanlanmış bir görev oluşturulur.

</details>

<details>
<summary><b>🖥️ Windows Server 2016'da Hangi Güncelleştirmeler Alınır?</b></summary>

Server 2016'da betik yalnızca otomatik güncelleştirmeleri `NoAutoUpdate` ilkesiyle durdurur; `wuauserv` hizmeti çalışmaya devam eder. Bu nedenle bilgisayar korumasız kalmaz.

ℹ️ **Bu, Windows Update'i durdurmanın en güvenli yöntemidir.** Hizmeti tamamen devre dışı bırakmak (`wuauserv` → Disabled), sistem dosyalarını silmek/yeniden adlandırmak veya üçüncü taraf "update blocker" araçları kullanmak gibi zorlayıcı yöntemlerin aksine; burada Microsoft'un **resmi, desteklenen ve geri alınabilir** Grup İlkesi ayarı (`NoAutoUpdate`) kullanılır. Hizmet çalışır durumda kaldığı için Defender güncel kalır, sistem kararlılığı bozulmaz ve devam ettirme işlemi tek tıkla, kalıcı bir hasar bırakmadan yapılır.

✅ **Duraklatmaya rağmen alınmaya devam edenler:**

| Güncelleştirme türü | Neden devam eder |
|---|---|
| Microsoft Defender virüs tanımları (imza / definition) | Defender bunları kendi mekanizmasıyla çeker (kendi zamanlanmış görevi, `MpCmdRun` / `Update-MpSignature`, WU Agent API ve MMPC doğrudan indirme yedeği). AU ilkesi bu akışı kapsamaz, böylece antivirüs koruması güncel kalır. |
| Elle başlatılan güncelleştirmeler | Bir yönetici "Güncelleştirmeleri denetle" derse yine de gelir (gerçek "pause" davranışıyla aynı). |

⛔ **Süre boyunca otomatik alınmayanlar (ertelenir, silinmez):**

| Güncelleştirme türü |
|---|
| Aylık toplu kalite/güvenlik güncelleştirmeleri (Cumulative Update / LCU) |
| Servicing Stack Update (SSU) ve .NET güncelleştirmeleri (WU üzerinden gelenler) |
| Sürücü güncelleştirmeleri (Windows Update kaynaklı) |
| Defender platform/motor aylık paketi (imzalardan farklıdır; günlük imzalar akmaya devam eder) |
| Özellik / sürüm yükseltmeleri |

Bu güncelleştirmeler silinmez, yalnızca bekletilir. Süre dolunca (zamanlanmış görev `NoAutoUpdate=0` yapar) veya betiği tekrar çalıştırıp devam ettirdiğinizde birikmiş güncelleştirmeler normal şekilde gelir.

⚠️ **Önemli uyarılar:**
- Etki alanına (domain) üye sunucularda, alan denetleyicisinden gelen bir Grup İlkesi Otomatik Güncelleştirmeleri ayarlıyorsa, **domain GPO yerel `NoAutoUpdate` ayarını ezebilir** ve duraklatma etkisiz kalabilir.
- **WSUS** kullanılıyorsa onay ve zamanlama oradan yönetildiğinden davranış farklılaşabilir.

</details>

<details>
<summary><b>📥 Kurulum ve Kullanım</b></summary>

1. Bu depo içerisindeki **`Windows_Update_Pause.bat`** dosyasını indirin. (Tek dosya çift dillidir: işletim sistemi diline göre Türkçe veya İngilizce çalışır.)
2. Dosyaya çift tıklayarak betiği çalıştırın (yönetici izni otomatik istenir, onaylayın).
3. Betik çalıştıktan sonra, güncelleştirmeleri kaç hafta duraklatmak istediğinizi belirtin (**1 - 10000 hafta arası**); belirttiğiniz süre boyunca güncelleştirmeler duraklatılır.
4. Güncelleştirmeleri devam ettirmek için betiği bir kere daha çalıştırmanız yeterlidir (betik, duraklatma durumunu algılayıp devam ettirir).

</details>

<details>
<summary><b>🧩 Desteklenen Sürümler</b></summary>

Betik, derleme numarasına göre (15063 eşiği) uygun yöntemi otomatik seçer.

✅ **Tam desteklenen:**

| Sürüm | Derleme | Kullanılan yöntem |
|---|---|---|
| Windows 11 (tümü, 24H2 / 26H1 dahil) | 22000+ | Yerleşik Pause API |
| Windows 10 1703 ve üzeri | 15063+ | Yerleşik Pause API |
| Windows 10 1507 / 1511 / 1607 | 10240–14393 | `NoAutoUpdate` ilkesi + görev |
| Windows Server 2022 | 20348 | Yerleşik Pause API |
| Windows Server 2019 | 17763 | Yerleşik Pause API |
| Windows Server 2016 | 14393 | `NoAutoUpdate` ilkesi + görev |
| Windows 8.1 / Server 2012 R2 | 9600 | `NoAutoUpdate` ilkesi + görev |
| Windows 8 / Server 2012 | 9200 | `NoAutoUpdate` ilkesi + görev |

⚠️ **Koşullu:** Windows 7 / Server 2008 R2 varsayılan olarak PowerShell 2.0 ile gelir ve gerekli cmdlet'ler (`Get-CimInstance`, `ScheduledTasks` modülü) bulunmadığından çalışmaz; yalnızca WMF 5.1 (PowerShell 5.1) elle kurulmuşsa çalışır.

⛔ **Desteklenmeyen:** Windows Vista / XP / Server 2008 ve öncesi.

</details>

<details>
<summary><b>🛠️ Sistem Gereksinimleri</b></summary>

- Windows 8 / Windows Server 2012 ve üzeri (Windows 8.1, 10, 11 ve Server 2012 R2, 2016, 2019, 2022 dahil)
- Windows PowerShell 3.0 veya üzeri (Windows 8 / Server 2012 ile birlikte gelir; Windows 10/11 ve Server 2016+ üzerinde 5.1)
- Yönetici yetkileri (betik dosyası tarafından otomatik sağlanır)

</details>

---

## 🇬🇧 English Explanation

<details>
<summary><b>🆕 What's New</b></summary>

- ⚙️ **PowerShell instead of VBS** — the code has been fully ported from `VBScript` to `PowerShell`. With a special **hybrid (polyglot) method**, the PowerShell code runs inside a single `.bat` file: the batch part at the top reads the file's own content and passes it to PowerShell (`Invoke-Expression`). This way it runs on double-click without dealing with `.ps1` execution policy.
- ✅ **Windows 10 / 11 support (incl. 26H1)** — uses the built-in *Pause Updates* API, writing `Pause*StartTime` / `Pause*ExpiryTime` with a real **UTC expiry date**, so updates are genuinely paused.
- 🖥️ **Windows Server 2016 and older support** — these versions have no Pause API; the script detects this automatically and stops automatic OS updates via the **`NoAutoUpdate` policy**. **The `wuauserv` service keeps running.** A **scheduled task** is created to turn the policy back off after the chosen number of weeks.
- 🛡️ **Security stays on** — with both methods only OS updates are stopped; **Microsoft Defender keeps receiving virus definition (signature) updates**, so your PC is not left unprotected.
- 🌐 **Automatic Turkish / English UI** — Turkish if the OS display language is Turkish, otherwise English.
- 📦 **Single file** — no separate TR/ENG versions; one script handles both languages.
- 🪟 **Console title** — the window title shows `Windows Update Pause / Resume | by Abdullah ERTÜRK`.

</details>

<details>
<summary><b>✨ Features</b></summary>

- **Easy to Use:** You can pause Windows updates with just a few double clicks.
- **Fast and Effective:** Pauses updates instantly, without wasting time.
- **Flexible Duration:** You choose the pause length in weeks (**between 1 and 10000 weeks**); invalid or out-of-range entries are rejected.
- **Reversible:** You can continue the updates from where you left off by running the script again.
- **Secure:** Operates securely using Windows' built-in commands and policies.
- **Defender Protected:** Even while updates are paused, Windows Defender keeps updating its virus definition (signature) database.
- **Broad Compatibility:** Supports Windows 10, Windows 11 and Windows Server 2016 (and older); the script detects the version and picks the best method automatically.
- **Automatic Language:** The script detects your OS language and shows the interface in Turkish or English.

</details>

<details>
<summary><b>⚙️ How It Works</b></summary>

- **Modern Windows 10 / 11 (build 15063 and above):** uses Windows' built-in "Pause Updates" feature; updates resume automatically when the period ends.
- **Windows Server 2016 and older (below build 15063):** since the Pause API is unavailable, automatic updates are stopped via the `NoAutoUpdate` policy, the `wuauserv` service keeps running, and a scheduled task is created to turn the policy back off when your chosen period ends.

</details>

<details>
<summary><b>🖥️ Which Updates Does Windows Server 2016 Receive?</b></summary>

On Server 2016 the script only stops automatic updates via the `NoAutoUpdate` policy; the `wuauserv` service keeps running, so the machine is not left unprotected.

ℹ️ **This is the safest way to stop Windows Update.** Unlike forceful methods such as fully disabling the service (`wuauserv` → Disabled), deleting/renaming system files, or using third-party "update blocker" tools, this uses Microsoft's **official, supported and reversible** Group Policy setting (`NoAutoUpdate`). Because the service stays running, Defender remains up to date, system stability is preserved, and resuming is a one-click action that leaves no lasting damage.

✅ **Still received even while paused:**

| Update type | Why it keeps working |
|---|---|
| Microsoft Defender virus definitions (signatures) | Defender pulls these through its own mechanism (its own scheduled task, `MpCmdRun` / `Update-MpSignature`, the WU Agent API and the MMPC direct-download fallback). The AU policy does not cover this flow, so antivirus protection stays current. |
| Manually triggered updates | If an admin clicks "Check for updates" they still arrive (same as the real "pause" behavior). |

⛔ **Not installed automatically during the period (deferred, not deleted):**

| Update type |
|---|
| Monthly cumulative quality/security updates (Cumulative Update / LCU) |
| Servicing Stack Updates (SSU) and .NET updates (delivered via WU) |
| Driver updates (sourced from Windows Update) |
| Defender platform/engine monthly package (different from signatures; daily signatures keep flowing) |
| Feature / version upgrades |

These updates are not deleted, only held back. When the period ends (the scheduled task sets `NoAutoUpdate=0`) or you run the script again to resume, the pending updates arrive normally.

⚠️ **Important notes:**
- On domain-joined servers, if a Group Policy from the domain controller configures Automatic Updates, the **domain GPO can override the local `NoAutoUpdate` setting** and the pause may have no effect.
- If **WSUS** is used, approval and scheduling are managed there, so the behavior may differ.

</details>

<details>
<summary><b>📥 Installation and Use</b></summary>

1. Download the **`Windows_Update_Pause.bat`** file in this repository. (The single file is bilingual: it runs in Turkish or English depending on the OS language.)
2. Run the script by double-clicking on the file (administrator permission is requested automatically; approve it).
3. After the script runs, specify the number of weeks you want to pause updates (**between 1 and 10000 weeks**); updates will be paused for that period.
4. To continue the updates, just run the script once more (it detects the paused state and resumes).

</details>

<details>
<summary><b>🧩 Supported Versions</b></summary>

The script automatically picks the right method based on the build number (the 15063 threshold).

✅ **Fully supported:**

| Version | Build | Method used |
|---|---|---|
| Windows 11 (all, incl. 24H2 / 26H1) | 22000+ | Built-in Pause API |
| Windows 10 1703 and later | 15063+ | Built-in Pause API |
| Windows 10 1507 / 1511 / 1607 | 10240–14393 | `NoAutoUpdate` policy + task |
| Windows Server 2022 | 20348 | Built-in Pause API |
| Windows Server 2019 | 17763 | Built-in Pause API |
| Windows Server 2016 | 14393 | `NoAutoUpdate` policy + task |
| Windows 8.1 / Server 2012 R2 | 9600 | `NoAutoUpdate` policy + task |
| Windows 8 / Server 2012 | 9200 | `NoAutoUpdate` policy + task |

⚠️ **Conditional:** Windows 7 / Server 2008 R2 ships with PowerShell 2.0 by default and will not work because the required cmdlets (`Get-CimInstance`, the `ScheduledTasks` module) are missing; it works only if WMF 5.1 (PowerShell 5.1) is installed manually.

⛔ **Not supported:** Windows Vista / XP / Server 2008 and earlier.

</details>

<details>
<summary><b>🛠️ System Requirements</b></summary>

- Windows 8 / Windows Server 2012 and later (including Windows 8.1, 10, 11 and Server 2012 R2, 2016, 2019, 2022)
- Windows PowerShell 3.0 or later (ships with Windows 8 / Server 2012; 5.1 on Windows 10/11 and Server 2016+)
- Administrator rights (provided automatically by the script file)

</details>
