Причина за грешката:
В GitHub са смесени scraper.py от v6 и tests/test_scraper.py от v5.
Старият тест очаква "Вътрешна дълбочина" да бъде външен размер, а v6 правилно я игнорира.

Замени в репото тези файлове със съдържанието на този архив:
- scraper.py
- tests/test_scraper.py
- .github/workflows/scrape.yml
- VERSION.txt

Преди качване изтрий стария tests/test_scraper.py и евентуални дублиращи файлове test_scraper (1).py.
След това пусни workflow в режим test.
Очакван резултат: Ran 32 tests ... OK.
