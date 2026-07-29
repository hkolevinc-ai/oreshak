import unittest
from decimal import Decimal
from bs4 import BeautifulSoup

import scraper


class ParserTests(unittest.TestCase):
    def test_price_is_taken_from_main_product_not_related_cards(self):
        html = '''
        <div id="content">
          <div class="product-info">
            <h1>ДЪСКА ЗА ТРОФЕЙ</h1>
            <ul class="list-unstyled"><li>Код на продукта: 5709</li></ul>
            <div class="price"><span class="price-old">80.00 €</span><span class="price-new">74.14 €</span></div>
            <button id="button-cart">Купи</button>
          </div>
          <div class="related"><div class="product-thumb"><h2>35.11 €</h2></div></div>
        </div>
        '''
        soup = BeautifulSoup(html, "lxml")
        price, list_price, source, list_source = scraper.OreshakClient._parse_prices(soup, {})
        self.assertEqual(price, Decimal("74.14"))
        self.assertEqual(list_price, Decimal("80.00"))
        self.assertIn("price-new", source)
        self.assertIn("price-old", list_source)

    def test_regular_price_not_minimum_price_on_page(self):
        html = '''
        <div id="content">
          <div class="product-info">
            <h1>ДЪРВОРЕЗБА ЗА ТРОФЕЙ ОТ ЕЛЕН 68</h1>
            <ul class="list-unstyled"><li>Код на продукта: 5314</li><li><h2>127.82 € (250.00 лв.)</h2></li></ul>
            <button id="button-cart">Купи</button>
          </div>
          <div class="product-grid"><h2>35.11 €</h2></div>
        </div>
        '''
        soup = BeautifulSoup(html, "lxml")
        price, list_price, _, _ = scraper.OreshakClient._parse_prices(soup, {})
        self.assertEqual(price, Decimal("127.82"))
        self.assertEqual(list_price, Decimal("127.82"))

    def test_json_ld_selects_matching_product(self):
        html = '''
        <h1>Main Product</h1>
        <script type="application/ld+json">[
          {"@type":"Product","name":"Related Product","url":"https://oreshak.bg/related","offers":{"price":"35.11","priceCurrency":"EUR"}},
          {"@type":"Product","name":"Main Product","url":"https://oreshak.bg/main","offers":{"price":"74.14","priceCurrency":"EUR"}}
        ]</script>
        '''
        soup = BeautifulSoup(html, "lxml")
        item = scraper.OreshakClient._json_ld_product(soup, "https://oreshak.bg/main", "Main Product")
        self.assertEqual(item["name"], "Main Product")

    def test_dimensions_with_labeled_object_and_decimal(self):
        text = "Размери на дъската: 35/23 см., Дебелина на дъската: 2 см., Тегло: 0,5 кг."
        self.assertEqual(scraper.parse_dimensions(text), (35.0, 23.0, 2.0))
        self.assertEqual(scraper.parse_weight(text), 500.0)

    def test_dimensions_repairs_split_number(self):
        text = "Размери на дъската : 3 9/29 см., Дебелина на дъската: 2 см."
        self.assertEqual(scraper.parse_dimensions(text), (39.0, 29.0, 2.0))

    def test_image_derivatives_are_deduplicated_and_spaces_encoded(self):
        html = '''
        <div id="content"><div class="product-info"><ul class="thumbnails">
          <li><a class="thumbnail" href="https://oreshak.bg/image/catalog/a/one image-1000x1000-product_popup.jpg"><img src="https://oreshak.bg/image/catalog/a/one image-500x500-product_thumb.jpg"></a></li>
        </ul></div></div>
        '''
        soup = BeautifulSoup(html, "lxml")
        images = scraper.OreshakClient._parse_images(soup, {}, "https://oreshak.bg/product")
        self.assertEqual(len(images), 1)
        self.assertIn("1000x1000-product_popup", images[0])
        self.assertNotIn(" ", images[0])

    def test_attributes_keep_decimal_values(self):
        description = "Описание. ➔ Характеристики: Размери на дъската: 31/30.5 см., Дебелина на дъската: 2 см., Тегло: 2.86 кг. ➔ Предимства: Ръчна изработка."
        attrs = scraper.OreshakClient._parse_attributes(BeautifulSoup("<div></div>", "lxml"), description)
        self.assertEqual(attrs["Размери"], "31/30.5 см")
        self.assertEqual(attrs["Тегло"], "2.86 кг")
        self.assertIn("30.5", attrs["Характеристики"])

    def test_stock_ignores_related_out_of_stock_text(self):
        html = '''
        <div id="content">
          <div class="product-info"><h1>Main</h1><button id="button-cart">Купи</button></div>
          <div class="related">Неналичен</div>
        </div>
        '''
        soup = BeautifulSoup(html, "lxml")
        available, source = scraper.OreshakClient._parse_stock(soup, {})
        self.assertTrue(available)
        self.assertIn("add-to-cart", source)


if __name__ == "__main__":
    unittest.main()
