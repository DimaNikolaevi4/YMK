# -*- coding: utf-8 -*-
"""ищи-3: скачивание перспективных источников (Полуянович Лань, Акимова, Келим, Воробьев, Селевцов, Шишмарёв trial)."""
import subprocess, os

OUT = '/home/z/my-project/YMK/docs/search/ishi3/fetch'
os.makedirs(OUT, exist_ok=True)

JOBS = {
    # Полуянович Лань — dist.berpt.ru папка с PDF (ОИ 4 КТП 05.02)
    'berpt_folder.html': 'https://dist.berpt.ru/mod/folder/view.php?id=165',
    # Акимова — полный PDF на elec.ru (ДИ 2 КТП 03.02)
    'akimova_elec.pdf': 'https://www.elec.ru/files/2019/10/16/%D0%9C%D0%BE%D0%BD%D1%82%D0%B0%D0%B6__%D1%82%D0%B5%D1%85%D0%BD%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B0%D1%8F_%D1%8D%D0%BA%D1%81%D0%BF%D0%BB%D1%83%D0%B0%D1%82%D0%B0%D1%86%D0%B8%D1%8F_%D0%B8_%D1%80%D0%B5%D0%BC%D0%BE%D0%BD%D1%82.PDF',
    # Келим — полный PDF (ДИ 5 КТП 03.02)
    'kelim_kstu.pdf': 'http://elib.kstu.kz/fulltext/Skan/kelim_asu.pdf',
    # Селевцов — scribd (частичное оглавление)
    'selivtsov_scribd.html': 'https://ru.scribd.com/document/990413298/%D0%A1%D0%B5%D0%BB%D0%B5%D0%B2%D1%86%D0%BE%D0%B2-%D0%9B-%D0%98-%D0%90%D0%B2%D1%82%D0%BE%D0%BC%D0%B0%D1%82%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D1%8F-%D0%A2%D0%B5%D1%85%D0%BD%D0%BE%D0%BB%D0%BE%D0%B3%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D1%85-%D0%9F%D1%80%D0%BE%D1%86%D0%B5%D1%81%D1%81%D0%BE%D0%B2',
    # Воробьев — studmed (ОИ 1 КТП 03.02)
    'vorobev_studmed.html': 'https://www.studmed.ru/vorobev-v-a-ekspluataciya-i-remont-elektrooborudovaniya-i-sredstv-avtomatizacii_d31a7f8983e.html',
    # Воробьев — dokumen.pub (возможно полный текст)
    'vorobev_dokumen.html': 'https://dokumen.pub/97d015d739a187555b13d9a9a613dae4.html',
    # Клепиков — dokumen.pub
    'klepikov_dokumen.html': 'https://dokumen.pub/2c81f87c901c77a5f520776e05f17c4b.html',
    # Шишмарёв Автоматика — trial PDF Литрес (обычно содержит оглавление)
    'shishmarev_avtomatika_trial.pdf': 'https://www.litres.ru/get_pdf_trial/25725613.pdf',
}

UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
for name, url in JOBS.items():
    dst = os.path.join(OUT, name)
    if os.path.exists(dst) and os.path.getsize(dst) > 1000:
        print(f'{name}: уже есть ({os.path.getsize(dst)} Б)')
        continue
    r = subprocess.run(['curl', '-sL', '--max-time', '60', '-A', UA, '-o', dst, '-w', '%{http_code} %{size_download}', url],
                       capture_output=True, text=True, timeout=90)
    print(f'{name}: HTTP {r.stdout.strip()}')
