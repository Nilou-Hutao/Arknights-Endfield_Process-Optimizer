<div align="right">
  <a href="#korean">🇰🇷 한국어</a> | <a href="#english">🇺🇸 English</a>
</div>

<a id="korean"></a>
# 🏭 명일방주: 엔드필드 공정 DB 및 최적화 계산기

엑셀 기반으로 제작된 엔드필드 스마트 공장 설계 및 레시피 검색 툴입니다. 아이템 이름만 검색하면 하위 공정부터 전력 소비량까지 한 번에 계산해 줍니다.

## ⚠️ 필수 사용 환경
* **엑셀 2013 이상 버전**에서만 정상 작동을 보장합니다.
* 본 파일은 VBA(매크로)가 포함된 `.xlsm` 파일입니다. 파일 실행 시 상단에 표시되는 **'콘텐츠 사용'** 또는 **'매크로 사용 방지 풀기'**를 반드시 클릭하셔야 정상 작동합니다.

---

## 🔍 1. 기본 검색 (생산/소모처 찾기)
특정 아이템이 어디서 만들어지고 어디에 쓰이는지 단순 검색할 때 사용합니다.

1. `'검색'` 시트를 엽니다.
2. `B2` 셀(*'검색할 텍스트 입력:'* 바로 옆 칸)에 찾으려는 아이템 이름을 입력합니다.
3. 바로 옆 **돋보기 아이콘**을 클릭합니다.
4. 생산 품목이나 소모 재료에 해당 검색어가 포함된 공정 목록이 출력됩니다.

---

## 🛠️ 2. 레시피 탐색 (최적화 공정 계산)
이 툴의 핵심 기능입니다. 특정 아이템 생산 공정이 멈추지 않고 계속 돌아갈 수 있도록 최소치 생산 공정을 계산해 줍니다.

1. 분석할 아이템 이름이 적힌 **셀을 클릭하여 선택**합니다.
2. 상단 중앙(`E2` 셀 우측)에 있는 **돋보기 아이콘**을 클릭합니다.

### 📊 검색 결과 항목 설명
* **총 전력소비량 (`G2` 셀):** 자원 채굴부터 최종 시설까지 전체 공정에서 소비되는 전력량의 합계입니다. 새로운 검색을 실행하면 자동으로 초기화됩니다.
* **공정 단계 (Tier):** 자원 채굴이 0티어이며, 생산 시설을 하나 거칠 때마다 1씩 증가합니다. 동일한 아이템을 만드는 공정이 여러 개일 경우, 프로그램이 자동으로 가장 티어가 낮은(효율적인) 공정을 우선하여 표시합니다.
* **소모 재료 관계:** `소모재료1(소비량) + 소모재료2(소비량) -> 생산 품목(생산량)` 형식으로 표기됩니다.
  * 재료가 1개라면 `소모재료(소비량) -> 생산 품목(생산량)`으로 표시됩니다.
  * 소비량이 `(0)`으로 표시되면 채굴해도 줄어들지 않는 무한 자원입니다.
* **잔여재료:** 공정 비율상 필연적으로 남게 되는 잉여 재료입니다.
  * 참고로 **분당 30개**는 컨베이어 벨트 1개로 온전히 옮길 수 있는 양(2초에 1개)을 의미합니다.
  * 파이프(액체)는 1초에 2개를 운송합니다.

---

## 💡 3. 주의사항 및 참고 팁
* 데이터 원본 시트는 수정하지 마시고, 가급적 `'검색'` 시트에서 검색어만 변경하며 사용하시길 권장합니다.
* 검색 결과를 띄워둔 상태로 엑셀을 저장하면 다음 실행 시에도 결과가 유지되며, 저장하지 않으면 이전 상태로 돌아갑니다.
* 돋보기 아이콘을 클릭한 뒤 최적화 경로를 계산하느라 **수 초 정도 딜레이**가 발생할 수 있으며, 이는 정상적인 처리 과정입니다.
* 검색을 반복하다 보면 엑셀 표 서식이 조금 깨질 수 있으나, 데이터를 확인하는 데는 지장이 없습니다.

---

## 💬 건의사항
추가되었으면 하는 기능이나 데이터상 잘못된 부분이 있다면 깃허브 **Issues** 탭에 남겨주시거나 커뮤니티 댓글로 편하게 말씀해 주세요. 
개인적인 일정으로 인해 즉각적인 반영은 어려울 수 있지만, 버전 업데이트 시 최대한 모아서 반영할 예정입니다.
<br>
<hr>
<br>

<a id="english"></a>
# 🏭 Arknights: Endfield Process DB & Optimization Calculator

An Excel-based smart factory planning and recipe search tool for Arknights: Endfield. Simply search for an item name to instantly calculate the entire production chain, from raw materials to total power consumption.

## ⚠️ System Requirements
* Guaranteed to work on **Microsoft Excel 2013 and later versions**.
* This file is an `.xlsm` format containing VBA macros. You must click **"Enable Content"** or **"Enable Macros"** at the top of the screen when opening the file for the tool to function correctly.

---

## 🔍 1. Basic Search (Find Production / Consumption)
Used to quickly find where a specific item is produced or consumed.

1. Open the `'검색'` (Search) sheet.
2. Enter the item name in cell `B2` (next to *'검색할 텍스트 입력:'*).
3. Click the **Magnifying Glass Icon** right next to it.
4. A list of all processes producing or consuming that item will be displayed.

---

## 🛠️ 2. Recipe Search (Process Optimization)
The core feature of this tool. It calculates the minimum required production facilities to keep the process running continuously without stopping.

1. **Click and select the cell** containing the item name you want to analyze.
2. Click the **Magnifying Glass Icon** located at the top center (next to cell `E2`).

### 📊 Search Result Explanations
* **Total Power Consumption (Cell `G2`):** The total sum of power required for the entire process line, from resource mining to the final facility. This automatically resets when a new search is performed.
* **Process Stage (Tier):** Resource mining starts at Tier 0, and increases by 1 for each production facility it passes through. If there are multiple ways to craft an item, the program automatically prioritizes the lowest tier (most efficient) process.
* **Material Relationship:** Displayed as `Material 1 (Consumption) + Material 2 (Consumption) -> Output Item (Production)`.
  * If there is only one material, it shows as `Material (Consumption) -> Output (Production)`.
  * If the consumption is shown as `(0)`, it indicates an infinite resource node that does not deplete.
* **Surplus Material:** Excess materials that inevitably remain due to process ratios.
  * *Note:* **30 per minute** is the exact amount to fully saturate one conveyor belt (1 item per 2 seconds).
  * Pipes (liquids) transport 2 units per second.

---

## 💡 3. Notes & Tips
* Please do not modify the raw data sheets. It is recommended to only change the search terms in the `'검색'` (Search) sheet.
* If you save the Excel file with the search results on screen, they will be kept for your next session. If you close without saving, it reverts to the previous state.
* There may be a **delay of a few seconds** after clicking the magnifying glass icon while it calculates the optimal path. This is completely normal.
* Repeated searches might cause the Excel table formatting (borders, colors) to misalign slightly, but it will not affect your ability to read the data.

---

## 💬 Feedback & Suggestions
If you have feature requests or spot any incorrect data, please feel free to leave a comment or open a ticket in the GitHub **Issues** tab. 
Due to my personal schedule, immediate fixes might be difficult, but I will do my best to compile and address them in future version updates.
