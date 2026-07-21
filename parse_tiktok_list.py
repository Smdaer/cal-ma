#!/usr/bin/env python3
"""Parse 'O bir defineci' TikTok video list into Excel."""

import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

RAW = r"""
KURD UP adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #anılar 1186
Sevinç🦋 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
1415
BABALARIN BABASI MÜSLÜM BABA adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan The support is unreal! Hit follow, join the fun #tiktoklive #livehighlights 
1078
can 🌹 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2466
Burak_music44 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #anılar 
6781
Mem Ararat adlı sanatçının Xaçirêk parçasıyla O bir defineci tarafından oluşturulan #keşfet 
5713
⚔️JesusAs ⚔️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
6696
Mem Ararat adlı sanatçının Xaçirêk parçasıyla O bir defineci tarafından oluşturulan #anılar 
3473
delildilanarofficial adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
32K
ÇORUMLU HAKAN adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2593
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
1788
Yusuf Can Türkmen adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
13.1K
ℍ𝕀ℙ𝔼ℝ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #trend #onthisday #parati #greenscreen #love 
3485
KURD UP adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #anılar 
12.8K
Bavê Can adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #anılar 
6095
muzik.kurdi1 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
5020
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2765
SİDRE adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
3474
ᴋᴜ̈ʀᴅ ᴍᴀᴍɪ̇ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
23.3K
🅑🅞🅩🅞 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #viraltiktok #keşfet 
8370
delildilanarofficial adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
3175
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2621
Fettah Güven Diyadin 04 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2541
Cemşi 21 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2841
🎶MUHAMMET ARAS🎶FAN adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #love #greenscreen #viraltiktok #öneal 
44.5K
✹𓃵MͩEͤZͮOPͩOͤTͮAͬMͥYͫA✹𓃵 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
5058
𝔻𝕖𝕣𝕤𝕚𝕞𝕝𝕚 𝕞𝕦𝕣𝕒𝕥 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2555
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
8144
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
3488
Romence pirens adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2421
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
32.6K
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
2971
👑Ali.👑 adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #anime #trend #explore #fypシ゚ 
3189
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
4348
Romence pirens adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
5157
DersimTV adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #kf #dance #humor #keşfet #xyzbca 
2447
Muzîkên Kurdî | Kurdish Songs adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar #defi 
6269
🥀zana birindare🥀 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
3018
Kadir Baş Erbaa Sokutaş lı 60 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılarım #keşfet #humor #dance#kf 
2617
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
30.7K
Uğur Şimşek adlı sanatçının Zilli Avrat & Kar Yolla parçasıyla O bir defineci tarafından oluşturulan #anılar #xyzbca #trend #dance #kf 
3681
Ayşe Şewaqî adlı sanatçının Pezkuvî parçasıyla O bir defineci tarafından oluşturulan #anılar #defineciler #anime #aaa #defi 
5542
🥀zana birindare🥀 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #kf #dance #trend #xyzbca 
6974
olgaamiyan adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #xyzbca #trend #dance #kf #keşfet 
6900
Ayşe Şewaqî adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar #dance #defender #defi #aaa 
5398
Zone Ma adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #kf #dance #trend #xyzbca 
2854
Ayşe Şewaqî adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #xyzbca #trend #dance #kf #keşfet 
3066
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #tik_tok #defineciler #trend #EF #mezmur 
2566
Yiğit İnsanların Türküleri adlı sanatçının Yiğit İnsanların Türküleri - orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #kf #dance #trend #xyzbca 
2741
Ɓeysu Ɓaşkan✌️📿 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
43K
illegalask_62 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #xyzbca #goviral #keşfet #humor #dance 
46.1K
Delîl Dîlanar adlı sanatçının Derweşe Evdi parçasıyla O bir defineci tarafından oluşturulan #trend #keşfet #humor #goviral #xyzbca 
2569
Kadir Baş Erbaa Sokutaş lı 60 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #keşfet #tik_tok #
4839
Bacımsu Bostan adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #xyzbca #şarkı #humor #goviral #dance 
17.1K
Bacımsu Bostan adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #dance #goviral #humor #şarkı #xyzbca 
3273
dr.faust09 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #f #tiktokviral #tiktokindia #trend #keşfet 
3517
Erdem Uygur adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar #trend #tiktokindia #tiktokviral #f 
2545
44.yüksel Arpacı.44 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar#f #trend #tiktokindia #tiktokviral 
3369
Yusuf Can Türkmen adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #tik_tok 
5392
İstanbul Sevdalısı.34.57 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #tik_tok 
7489
dr.faust09 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #anılar #keşfet 
3049
44.yüksel Arpacı.44 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
4255
𝑩𝒆𝒌𝒊𝒓 𝑴𝒖𝒔𝒕𝒂𝒇𝒂27 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #keşfetteyizzz #anime #CEL #efkar 
16K
Kaşanlı Emiş adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2373
HEWİ✌🏿🖤🥀 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
33.6K
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar
2971
Վան   Բիթլիս 𒆜𓊉꧂ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3959
Serkan Kanîreş adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
3031
𝑩𝒆𝒌𝒊𝒓 𝑴𝒖𝒔𝒕𝒂𝒇𝒂27 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2321
♥️hayat ne garip ane♥️♥️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
24.9K
kurdarchive adlı sanatçının kurdarchive - orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar 
74.8K
MISHA_EZDI adlı sanatçının оригинальный звук parçasıyla O bir defineci tarafından oluşturulan 
3899
akif.ozek adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2572
JaPoN MecNuR TopRaK adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2728
Musa Eroğlu adlı sanatçının İlme Değer Verdim parçasıyla O bir defineci tarafından oluşturulan 
3264
MISHA_EZDI adlı sanatçının оригинальный звук parçasıyla O bir defineci tarafından oluşturulan 
1985
Yemci Fesih Mamuncu adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2218
🎥🎵SERHAT✅️3️⃣6️⃣🎵🎥 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
9424
Cano İbo adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2854
Munzur🌹 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
5555
sudeelilonline adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
11.5K
ERA adlı sanatçının The Mass parçasıyla O bir defineci tarafından oluşturulan #keşfet 
1973
ÖMƏR..1990 adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3663
Darya Ebrahimi adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2094
Seçil Ezircan adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2219
Reso adlı sanatçının Zerdûşt parçasıyla O bir defineci tarafından oluşturulan Zerdüşt peygamber 
4273
user13763600524 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
30.4K
ERA adlı sanatçının Ameno - Remix parçasıyla O bir defineci tarafından oluşturulan #keşfet 
12.4K
Ömer amed75 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
6620
کاروان adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2745
TRT Müzik adlı sanatçının original sound - TRT Müzik parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3448
🎶ÇİSEM🎶61🇹🇷55 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
4193
renas yolbilen adlı sanatçının EZ DILGEŞE CANIM - COVER parçasıyla O bir defineci tarafından oluşturulan 
13.7K
1heval__ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2562
Redkit TV adlı sanatçının Redkit TV - orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar
6306
Rebecca Mia adlı sanatçının Yabancıyım Kendime parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz 
4901
Masiroj Produktion adlı sanatçının JekTV_international - orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kf #defisitkalori #mezmur #mezarcı #k #defineciler #defender #keşfetteyizzz #kefşetbeni #keşfet 
7427
🅨︎🅤︎🅝︎🅤︎🅢︎ 🅢︎İ🅟︎🅐︎🅝︎ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kf 
5245
misracglr adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kf #kefşetbeni #keşfet #keşfetteyizzz #defender #defineciler #k #mezarcı #mezmur #defisitkalori #kf 
21K
🅨︎🅤︎🅝︎🅤︎🅢︎ 🅢︎İ🅟︎🅐︎🅝︎ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kf 
28.4K
༄ 𝔸𝕝𝕒𝕟༄ adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #kf 
5505
MISHA_EZDI adlı sanatçının оригинальный звук parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineciler #k #mezarcı #mezmur #defisitkalori #defender #kf #keşfetteyizzz #kefşetbeni 
4153
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2515
e1yaz__1 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #defineciler #keşfet #k #mezarcı #mezmur #defisitkalori #defender #kf #keşfetteyizzz #kefşetbeni 
3785
T Ü R K Ü  S O K A Ğ I adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineciler #
5001
Mamostesezer adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet #kefşetbeni #defineci #defi #keşfet #efkar #CEL #anime 
14K
🇩🇪❤️جان  جان ❤️🇩🇪 adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #kefşetbeni #keşfet #keşfetteyizzz #kf #defender #defisitkalori #defineciler #mezarcı #mezmur 
2735
FLKLORE DEVERA BADINAN adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #kesfet #keşfet #defense #ef #deficitcalorico #defender #defense #defineci #defi #kefşetbeni #kefşet 
11.9K
𝐌𝐀𝐌𝐎 𝐀𝐅𝐑İ𝐍 366 🕊️ adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #kefetteyiz #efootballmobile #defineciler #kefşett #kesfetteyiz #define #definearama #mezarcı #dance #ef #keşfet 
6405
Meral Alkan adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #definearama #define #kesfetteyiz #kefşett #defineciler #kefetteyiz #efootballmobile 
44.1K
Fairuz adlı sanatçının El Massih Kaam parçasıyla O bir defineci tarafından oluşturulan #kefetteyiz #efootballmobile #defender #defisitkalori #ef #defineciler #kefşett #dance #kesfetteyiz #define #definearama #mezarcı 
6733
Fᴀʀɪᴅᴀ KURDÎ ⚘️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kesfetteyiz #dance #kefşett #defineciler #ef #defisitkalori #kefetteyiz #defender #efootballmobile 
12.3K
__pivok__ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kesfetteyiz #kefşet #keşfet #defender #define 
10.7K
denizersin36 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet #defineciler #keşfetteyizzz #
10.2K
🅱🆈.🅷🅰🅻🅸̇🅻💎 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kesfetteyiz #kefşet 
4098
Munzur adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet 
30.7K
Harun Tunç adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defisitkalori #kefetteyiz #defender #define #definearama #ef #defineciler #kefşett #kesfetteyiz #dance 
10.8K
Sevgi Nur Ayaz adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #dance #kesfetteyiz #kefşett #defineciler #ef #definearama #define #defender #kefetteyiz #defisitkalori 
6164
Denge Kurdi adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #define #defender #kefetteyiz #defisitkalori #definearama #ef #defineciler #kefşett #kesfetteyiz #dance 
50K
Levent Aktaş adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett #kesfetteyiz #defineciler #defisitkalori #ef #definearama #defisitkalori #kefetteyiz #defender #define 
75.3K
مصطفى صبري adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #kesfetteyiz 
4298
M Semih T adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet 
8579
Ian Post adlı sanatçının Yole Kross Ayle (The Song of the Dead) parçasıyla O bir defineci tarafından oluşturulan #kefşet 
7492
Hüseyin 6565 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett 
30.3K
Munzur adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett #defineciler #defisitkalori #ef #definearama #define #defender #kefetteyiz #keşfet 
10.2K
Şahin Çelik adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett 
5418
🍂Zeynep🍂 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefetteyiz #defender #define #definearama #kefşett #ef #defisitkalori #defineciler 
18.1K
Kaos Production adlı sanatçının Ex Production Serayê KURDİSH REMİX parçasıyla O bir defineci tarafından oluşturulan #kefşett #ef #kefetteyiz #definearama #define #defender #kefşetbeni 
3360
Hewi🌙 ރ⁶⁵ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet 
5094
Anonim Muhabir adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett 
4395
user44660778002 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
5641
bahozbakur1 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett 
8979
Efrîna min adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #define #definearama #kefetteyiz #ef #kefşett 
33.7K
Hewi🌙 ރ⁶⁵ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
7418
Sezai adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet 
22.9K
adelafrini adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #define #ef #defineciler #kefetteyiz #keşfet #mezopotamya #definearama 
9414
"مهستی" adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #defense #define #kefşet #ef #defisitkalori #deficitcalorico 
13.7K
Mehdî Mêhvane adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfet #kefetteyiz #define #defineciler #ef 
22.7K
user44660778002 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defisitkalori #ef #kefşet #define #defense #keşfetteyizzz 
4809
Rojda şenses adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
48K
ÇETO MEDYA adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #define #define #kefetteyiz #defineciler 
12.4K
sinan koçak adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3670
ezeldemirel adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #define #keşfetteyizzz #kefşet #defisitkalori #ef 
25K
BAVE DILYAR💚☀️❤️ adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #defender #kefşetbeni #ef #defineciler #kefset #kefetteyiz #keşfet #define #mezopotamya 
9505
🎶EZDI🎶 MUSIC🎶 adlı sanatçının оригинальный звук parçasıyla O bir defineci tarafından oluşturulan #kefset #defineciler #ef #kefşetbeni #defender 
5702
user36360938076 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defender #kefşetbeni #ef #defineciler #kefset 
3190
user86991008439 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #defense #defisitkalori #kefşet #keşfetteyizzz #define 
3422
CîRO adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefset #defineciler #ef #kefşetbeni #defender 
2970
user73391075811 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefset #defineciler #kefşetbeni #ef 
29.8K
Fᴀʀɪᴅᴀ KURDÎ ⚘️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #kefset 
4375
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #kefetteyiz #kefşetbeni #kefset #defineciler 
3863
Ez kesek im ku di êşê de dijî adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #kefset #kefşetbeni #kefetteyiz 
7046
Cano İbo adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #define #kefseteyiz #kefset #defineciler 
7243
Fᴀʀɪᴅᴀ KURDÎ ⚘️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kesfet #defineciler 
2738
MUHAMMED👑👑👑 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kesfet #
2845
𝓒𝓪𝓷𝓮_47 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #defi #defender #keşfetteyizzz #keşfet #efkar #tik_tok 
5706
müşfiq_dağlar_oğlu adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2598
🌊 Kemal Sunal short 🌊 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefseteyiz #keşfet 
2960
ihsan YALVARMAZ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşetbeni #kefset #kefseteyiz #define 
23.1K
Keça Kobani adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet 
2348
Vian❤🦚🦚👑 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz 
7072
Irak Türkleri adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet 
2539
YERAZ🤘Qarabağ veteranı 🏅🎖️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefseteyiz #keşfet 
9939
𝑌𝑎𝑧𝑌𝑎ğ𝑚𝑢𝑟𝑢𝑚.. adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #kefseteyiz 
2604
BURHAN ÇETİN adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2679
jemilealyousuf(Pire) adlı sanatçının originalljud parçasıyla O bir defineci tarafından oluşturulan #keşfet 
5460
Alevi✌️im adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfet 
35.2K
Şeroyê Biro adlı sanatçının Desmala Min parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2481
Carlos Estella adlı sanatçının Flying Beyond the Sky parçasıyla O bir defineci tarafından oluşturulan #keşfet 
8955
🄲🄶🅂 🄾🄵🄵🄸🄲🄸🄰🄻 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2938
user49655852229 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
39.2K
👑 Emirhan Akkuş 👑 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
14.1K
Aram Tîgran adlı sanatçının Rıhe parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2574
YERAZ🤘Qarabağ veteranı 🏅🎖️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2344
Sehmus Temel🎼 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
6392
user25195426151 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #
27.8K
Hewidar adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
56.4K
user13017432290 adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #keşfet 
8068
Sivaslı kara adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #ef 
7808
Aram Tigran adlı sanatçının Gelo Ew Ki Bu ? parçasıyla O bir defineci tarafından oluşturulan #keşfet #keşfetteyizzz 
2742
kadir3065 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
6451
user02434698443 adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3353
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2485
DEMA GEL adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #keşfet 
10.1K
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineciler 
38.1K
DEMA GEL adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3317
CîRO adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #ef #keşfet 
2782
🤍🅴🆉🅸🅳🅸🆈🅰️ adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #ef #defineciler #keşfet 
17.6K
Suleyman Hüseyin💎 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #ef
13.2K
1 Dakika Kürtçe adlı sanatçının 1 Dakika Kürtçe - orijinal ses parçasıyla O bir defineci tarafından oluşturulan #ef 
22.4K
𝕯𝖊𝖗𝖝𝖔_12🖤 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
6472
mert adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #define #defineciler 
13.3K
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #anılar #definear #define #revenge #definefreshchallenge #efkarli1adam #kefşett #reviewanngon #efkar #defi #definefreshchallenge #yeziden #kefşetbeni #edit #keşfetteyizzz #doguinho #deficitcalorico 
11.5K
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #ermenistan #yeziden #kefşetbeni #edit #efkar #keşfetteyizzz #doguinho #deficitcalorico #defi #efkar #reviewanngon #kefşett #efkarli1adam #definefreshchallenge #revenge #define #definear 
4533
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşett 
2302
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #kefşett 
4296
04_Hewi_04 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #ef 
2483
Nizamettin Ariç adlı sanatçının Ay Dil parçasıyla O bir defineci tarafından oluşturulan #define #defi #ermenistan #mezopotamya #yeziden #kefşetbeni #edit #efkar #keşfetteyizzz #doguinho #efkarli1adam #keşfet #definefreshchallenge #revenge #reviewanngon 
5968
AZİZ TEK adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #mezopotamya#ermenistan #defi #define 
2844
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #kefşetbeni 
5544
💔zalimin kızı💔 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #mezopotamya #erivan #defineciler #kefşetbeni #feke #anılar #defisitkalori #keşfetteyizzz 
2727
urfa ve tarih adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #kur #mezarcı #mezopotamya #kurdishgirl #defi #defineciler #definefreshchallenge #kefşetbeni 
14.2K
TRT Kurdi TV adlı sanatçının TRT Kurdi TV - orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
14.4K
Ezdixan adlı sanatçının оригинальный звук parçasıyla O bir defineci tarafından oluşturulan #keşfet #mezopotamya #defense #definefreshchallenge 
2338
Şeroyê Biro adlı sanatçının Bedewê parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #erivan
2716
AZİZ TEK adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #keşfet #kefşetbeni #defineciler #keşfetteyizzz 
4483
kemal_sunalfilm adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #defisitkalori #keşfetteyizzz #kefşet #defineciler #define #anılar #definefreshchallenge #efkar #edit #feke #kefşetbeni #keşfet 
4520
Leyla🦚 adlı sanatçının originalljud parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #efkarli1adam #ef #definebulma #deficitcalorico #defisitkalori #define #kefşet #defineciler #keşfetteyizzz #efkar #kefşetbeni 
11K
Karadeniz Nağmeleri adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefetteyiz #efkar #keşfetteyizzz #doguinho #kefşet #define #defisitkalori #deficitcalorico #definebulma #ef #efkarli1adam #mezopotamya 
7803
Sami Yusuf adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #definefreshchallenge #defineciler #defisitkalori #kefşet #keşfetteyizzz #define #anılar #efkar #kefetteyiz #doguinho 
10.2K
merwan968 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #kefşet #defisitkalori #defineciler #definefreshchallenge 
34.2K
ابو رودي adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #kefşet #defisitkalori #defineciler #define #anılar #definefreshchallenge #efkar #edit #feke #kefetteyiz #mezopotamya 
35.4K
user98728770565 adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #mezopotamya #feke #edit #efkar #definefreshchallenge #anılar #define #defineciler #kefşet #keşfetteyizzz #defisitkalori 
37.4K
urfa ve tarih adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşetbeni 
29.3K
XECÊ HERDEM FAN👑 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefşet #kefşetteyizz #kefetteyiz #kefşetbeni 
28.4K
stranekurdi adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineciler 
2888
@Valateme adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
6107
Sehfe Satılır 💸 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #efkar #defineciler #keşfet #definefreshchallenge #define 
27.3K
MUZÎKA KURDÎ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #feke #edit #keşfet #keşfetteyizzz #defineciler #define #efkar #definefreshchallenge #defisitkalori 
20.6K
Kurtuluş❤DoğuKars İstanbul adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #definefreshchallenge #efkar #define #defineciler #keşfetteyizzz #keşfet #edit #feke 
15K
𝐁𝐞𝐫𝐳𝐚𝐧𝐱𝟒𝟗 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2591
EᖇEᑎ ᗪEᖇᔕİᗰ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet #define #defineciler 
2599
denizersin36 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #efkarli1adam #keşfetteyizzz #keşfet #defisitkalori #defineciler #define 
2080
Kurdish Music 21 adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #keşfet #keşfetteyizzz 
1751
Sarı Gül adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet #defineciler #ef 
1827
firdevsmusic adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #gomubulma #deficitcalorico #define #definebulma #defineciler #keşfet #keşfetteyizzz 
2065
mertTunc27 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #definebulma #define #deficitcalorico #gomubulma 
16.9K
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #deficitcalorico #define #definebulma #defineciler #defineciler 
7166
Irak Türkleri adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler 
1993
STORİLİK adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet #defineciler #tarihikentmardin #efkar 
2388
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #hayvansevgisi #defineciler #keşfet #keşfetteyizzz #hayvanlarıkoruyalım 
1783
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
1688
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet 
1959
Sami Yusuf adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #efkar #keşfet #keşfetteyizzz #defineciler 
2122
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #defineciler #keşfetteyizzz #keşfet #efkar 
12.5K
XECÊ HERDEM FAN👑 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #efk #keşfet #keşfetteyizzz #defineciler #define 
2382
All.Ezidi adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #fke #keşfet #efk 
13.5K
kocbertv adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #efk #keşfet #defi #defineciler #defisitkalori 
7522
Yusuf Roman adlı sanatçının Hayho parçasıyla O bir defineci tarafından oluşturulan #keşfet #definebulma #define 
26.1K
 adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #define #keşfet 
2804
Hewi🌙 ރ⁶⁵ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
2152
Hamitkaradeniz_official ☑️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet 
3096
Mir Perwer adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineci 
13.4K
Aşık Veysel adlı sanatçının Benim Sadık Yarim Kara Topraktır parçasıyla O bir defineci tarafından oluşturulan #kefseteyiz #tiktok 
21.1K
R.Ergenekon adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineci #keşfet #define #tiktok #definebulma #dogal #turkey 
3303
 adlı sanatçının original sound parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #defineci 
26.6K
Omid adlı sanatçının Hareeme Asheghi parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineci #tarihi 
2608
@Але_алехандро21 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #definebulma #defineci #kefseteyiz 
2180
⭐️ᴼᶻᴳᵁN.11⭐️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #kefsetbeniöneçıkar 
4302
STORİLİK adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #tiktok #defineci #keşfetteyizzz 
18.1K
Priya adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #keşfet #define 
21K
Kalbin1Ritim🖤 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefseteyizz 
8126
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan #keşfet #define #definebulma #defineci #turkey 
25.4K
Ebru kurt adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #kefseteyizz #defineci #definebulma #define #keşfet #turkiye 
3290
Maşallah Yalvarici adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineci 
4011
Diroka Ezdixane adlı sanatçının son original parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineci#define #turkiye 
20.2K
ciwan adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineci 
2396
All.Ezidi adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #defin #definebulma #defineci #keşfet 
9214
Mem Ararat adlı sanatçının Xaçirêk parçasıyla O bir defineci tarafından oluşturulan #turkey #keşfet #defineciler #dunyakobataado 
12.8K
Homeyra adlı sanatçının Dooreh Dooreh parçasıyla O bir defineci tarafından oluşturulan #defineciler #keşfetteyizzz #tektek #keşfet #dunyakobataado #turkiye 
29.2K
Fairuz adlı sanatçının El Massih Kaam parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #defineciler 
13.1K
Murat2434 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #define #defineciler #keşfetteyizzz #keşfet #doga #tarihi 
3208
Ayşe demir adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #keşfetteyizzz #defineciler #define #doga 
4893
🌼🌼🌼SEVILAY🌼🌼🌼 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler 
6824
Beşir Kaya adlı sanatçının Sallana Sallana parçasıyla O bir defineci tarafından oluşturulan #keşfetteyizzz #keşfet #defineciler #define 
3252
Ali Haydar Canbulat adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #defineciler #define işaretleri #keşfetteyizzz #tarihi eserler 
14.3K
Kibariye adlı sanatçının Allah Vergisi parçasıyla O bir defineci tarafından oluşturulan #defineciler #tarihi #keşfetteyizzz #define #keşfet 
2589
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #keşfet #keşfetteyizzz #define #defineciler #tarihi 
2478
drmuzikk adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #keşfet #keşfetteyizzz #define 
15.9K
Samir🇦🇿🇺🇸 adlı sanatçının оригинальный звук parçasıyla O bir defineci tarafından oluşturulan #keşfet #define #defineciler #keşfetteyizzz 
2717
alik adlı sanatçının Originalton parçasıyla O bir defineci tarafından oluşturulan #dogal #defineciler #define #keşfet #tilsmagning 
2501
munzurjiyan adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #define #defineciler #dogal #kefşet #tilsmagning 
21.7K
İbrahim kaya adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #defineciler #dogal #kefşet #dogandcat 
30.6K
Neriman Kayseri adlı sanatçının Kahretmişim Hayata parçasıyla O bir defineci tarafından oluşturulan #kefşet #defineciler 
7608
gülcan kaya adlı sanatçının Bende Şu Dünyaya Geldim Geleli parçasıyla O bir defineci tarafından oluşturulan #kefşet #doğavideolari #definevideoları 
17.8K
KURDÎ MUSÎC 💚☀️❤️ adlı sanatçının الصوت الأصلي parçasıyla O bir defineci tarafından oluşturulan #definevideoları #doğavideolari 
5239
pembee mezarlik✅ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
5279
🅱🆈.🅷🅰🅻🅸̇🅻💎 adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #definevideoları #turkiye #doğavideolari #mardin #bursa #diyarbakır 
17.7K
Ömer asaf özdemir adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #definevideoları #turkiye 
2904
İbrahim Tatlıses adlı sanatçının Yorgunum parçasıyla O bir defineci tarafından oluşturulan #turkiye #iznik #defilecik #definevideoları #mardin #doğavideolari#bursa   #diyarbakır 
3601
Karadeniz Akustik ♾️ adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan #mardin #bursa #definevideoları #DoğaVideolari #diyarbakır #iznik #turkiye 
3177
 adlı sanatçının  parçasıyla O bir defineci tarafından oluşturulan 
5637
Nilüfer Akbal adlı sanatçının Dilo Yeman parçasıyla O bir defineci tarafından oluşturulan 
2331
Arabesk Diyarı adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
2538
O bir defineci adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan 
3912
Kazım Koyuncu adlı sanatçının Gelevera Deresi (feat. Şevval Sam) parçasıyla O bir defineci tarafından oluşturulan 
2745
TRT Arşiv adlı sanatçının orijinal ses parçasıyla O bir defineci tarafından oluşturulan ##Turkiye 
2865
Anoushirvan Rohani adlı sanatçının Tarane Saal parçasıyla O bir defineci tarafından oluşturulan 
3154
Cavit Karabey adlı sanatçının Vay Be parçasıyla O bir defineci tarafından oluşturulan 
2687
Cemile Sönmez adlı sanatçının Meyrik (Uzun Hava) parçasıyla O bir defineci tarafından oluşturulan 
5734
Devran Çağlar adlı sanatçının Bir Kadın Yüzünden parçasıyla O bir defineci tarafından oluşturulan Mardine selam olsun . . . #fyp #diyarbakır #mardin #defilecik #iznik #bursa #Turkiye #keşfet 
5715
Ezgin Güzelgül adlı sanatçının Gelmiş Bahar Geçmiş Yazlar parçasıyla O bir defineci tarafından oluşturulan 
7010
Bergen adlı sanatçının Onu da Yak Tanrım parçasıyla O bir defineci tarafından oluşturulan 

3349
"""

ENTRY_RE = re.compile(
    r"^(?P<artist>.*?)\s*adlı sanatçının\s*(?P<sound>.*?)\s*parçasıyla\s+"
    r"O bir defineci tarafından oluşturulan\s*(?P<caption>.*?)\s*$",
    re.DOTALL,
)
COUNT_RE = re.compile(r"^(\d+(?:\.\d+)?K?)$", re.IGNORECASE)
TRAILING_COUNT_RE = re.compile(r"^(.*?)(?:\s+)(\d+(?:\.\d+)?K?)\s*$", re.IGNORECASE)


def parse_count(text: str) -> tuple[str, int | None]:
    text = text.strip()
    m = COUNT_RE.match(text)
    if not m:
        return text, None
    raw = m.group(1)
    if raw.upper().endswith("K"):
        num = float(raw[:-1]) * 1000
        return raw, int(num)
    return raw, int(raw)


def parse_entries(raw: str) -> list[dict]:
    lines = [ln.rstrip() for ln in raw.strip().splitlines()]
    entries: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # Accumulate multi-line until we have a full entry pattern or next entry starts
        block = line
        # If this line alone is just a count leftover, skip
        if COUNT_RE.match(line) and not entries:
            i += 1
            continue
        if COUNT_RE.match(line):
            # orphan count – attach to previous if missing
            if entries and not entries[-1].get("goruntulenme"):
                raw_c, val = parse_count(line)
                entries[-1]["goruntulenme"] = raw_c
                entries[-1]["goruntulenme_sayi"] = val
            i += 1
            continue

        # Peek next line for count-only
        caption = ""
        artist = sound = None
        m = ENTRY_RE.match(block)
        if not m:
            # Try combining with next non-empty if somehow split (shouldn't happen)
            i += 1
            continue

        artist = m.group("artist").strip()
        sound = m.group("sound").strip()
        caption = m.group("caption").strip()

        goruntulenme = ""
        goruntulenme_sayi = None

        # Count may be trailing on same line
        if caption:
            tm = TRAILING_COUNT_RE.match(caption)
            if tm and COUNT_RE.match(tm.group(2)):
                # Only treat as count if rest looks like hashtags/caption without digits at end of words oddly
                # and next line isn't also a count... prefer next-line count when present
                next_is_count = False
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    j += 1
                if j < len(lines) and COUNT_RE.match(lines[j].strip()):
                    next_is_count = True
                if not next_is_count:
                    caption = tm.group(1).strip()
                    goruntulenme, goruntulenme_sayi = parse_count(tm.group(2))

        # Or count on following line
        if not goruntulenme:
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and COUNT_RE.match(lines[j].strip()):
                goruntulenme, goruntulenme_sayi = parse_count(lines[j].strip())
                i = j  # consume count line

        entries.append(
            {
                "sanatci": artist,
                "parca": sound,
                "aciklama": caption,
                "goruntulenme": goruntulenme,
                "goruntulenme_sayi": goruntulenme_sayi,
            }
        )
        i += 1
    return entries


def write_excel(entries: list[dict], path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Videolar"

    headers = [
        "No",
        "Sanatçı / Ses Sahibi",
        "Parça / Ses Adı",
        "Açıklama / Hashtagler",
        "Görüntülenme",
        "Görüntülenme (Sayı)",
        "Oluşturan",
    ]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    alt_fill = PatternFill("solid", fgColor="F2F2F2")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(1, col, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin

    for idx, e in enumerate(entries, 1):
        row = [
            idx,
            e["sanatci"],
            e["parca"],
            e["aciklama"],
            e["goruntulenme"],
            e["goruntulenme_sayi"],
            "O bir defineci",
        ]
        for col, val in enumerate(row, 1):
            cell = ws.cell(idx + 1, col, val)
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if idx % 2 == 0:
                cell.fill = alt_fill

    widths = [6, 36, 42, 55, 14, 18, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.auto_filter.ref = f"A1:G{len(entries) + 1}"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28

    # Summary sheet
    ws2 = wb.create_sheet("Ozet")
    total = len(entries)
    with_count = sum(1 for e in entries if e["goruntulenme_sayi"] is not None)
    total_views = sum(e["goruntulenme_sayi"] or 0 for e in entries)
    top = sorted(
        [e for e in entries if e["goruntulenme_sayi"] is not None],
        key=lambda x: x["goruntulenme_sayi"],
        reverse=True,
    )[:20]

    ws2["A1"] = "Özet"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2["A3"] = "Toplam video"
    ws2["B3"] = total
    ws2["A4"] = "Görüntülenmesi olan"
    ws2["B4"] = with_count
    ws2["A5"] = "Toplam görüntülenme (yaklaşık)"
    ws2["B5"] = total_views

    ws2["A7"] = "En çok görüntülenen 20 video"
    ws2["A7"].font = Font(bold=True)
    for col, h in enumerate(["No", "Sanatçı", "Parça", "Görüntülenme", "Sayı"], 1):
        c = ws2.cell(8, col, h)
        c.font = header_font
        c.fill = header_fill

    for i, e in enumerate(top, 1):
        ws2.cell(8 + i, 1, i)
        ws2.cell(8 + i, 2, e["sanatci"])
        ws2.cell(8 + i, 3, e["parca"])
        ws2.cell(8 + i, 4, e["goruntulenme"])
        ws2.cell(8 + i, 5, e["goruntulenme_sayi"])

    for i, w in enumerate([6, 36, 42, 14, 12], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)


def write_csv(entries: list[dict], path: str, delimiter: str = ";") -> None:
    import csv
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    def clean(s):
        if s is None:
            return ""
        return ILLEGAL_CHARACTERS_RE.sub("", str(s)).replace("\n", " ").strip()

    headers = [
        "No",
        "Sanatci / Ses Sahibi",
        "Parca / Ses Adi",
        "Aciklama / Hashtagler",
        "Goruntulenme",
        "Goruntulenme (Sayi)",
        "Olusturan",
    ]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=delimiter)
        w.writerow(headers)
        for i, e in enumerate(entries, 1):
            w.writerow([
                i,
                clean(e["sanatci"]),
                clean(e["parca"]),
                clean(e["aciklama"]),
                clean(e["goruntulenme"]),
                e["goruntulenme_sayi"] if e["goruntulenme_sayi"] is not None else "",
                "O bir defineci",
            ])


def main() -> None:
    entries = parse_entries(RAW)
    out = "/workspace/O_bir_defineci_videolar.xlsx"
    write_excel(entries, out)
    write_csv(entries, "/workspace/O_bir_defineci_videolar.csv", ";")
    write_csv(entries, "/workspace/O_bir_defineci_videolar_virgullu.csv", ",")
    print(f"Toplam kayıt: {len(entries)}")
    missing = [i + 1 for i, e in enumerate(entries) if not e["goruntulenme"]]
    print(f"Görüntülenmesi eksik: {len(missing)} -> {missing[:10]}")
    for e in entries[:3]:
        print(e)
    print("...")
    for e in entries[-3:]:
        print(e)
    print(f"Kaydedildi: {out}")
    print("CSV: O_bir_defineci_videolar.csv")


if __name__ == "__main__":
    main()
