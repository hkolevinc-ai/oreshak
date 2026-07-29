import unittest
from decimal import Decimal
from bs4 import BeautifulSoup

import scraper


class ParserTests(unittest.TestCase):
    def test_project_version_matches_bundle(self):
        self.assertEqual(scraper.PROJECT_VERSION, "6.1")

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

    def test_hidden_checkout_price_beats_visible_old_price(self):
        html = '''
        <div id="content"><div class="product-info">
          <h1>КОМПЛЕКТ ШАХ И ТАБЛА</h1>
          <ul class="list-unstyled"><li><h2 class="price">81.81 €</h2></li></ul>
          <input id="price" name="price" type="hidden" value="77.72">
          <button id="button-cart">Купи</button>
        </div></div>
        '''
        soup = BeautifulSoup(html, "lxml")
        price, list_price, source, list_source = scraper.OreshakClient._parse_prices(soup, {})
        self.assertEqual(price, Decimal("77.72"))
        self.assertEqual(list_price, Decimal("81.81"))
        self.assertIn("input#price", source)
        self.assertIn("higher than checkout price", list_source)

    def test_regular_price_has_no_manufactured_list_price(self):
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
        self.assertIsNone(list_price)

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

    def test_outer_dimensions_beat_compartment_dimensions(self):
        text = "Външни размери: 25/15/10 см. Размери на едно отделение: 11.5/10.5/3 см."
        self.assertEqual(scraper.parse_dimensions(text), (25.0, 15.0, 10.0))

    def test_closed_dimensions_beat_square_size(self):
        text = "Размери в затворено състояние: 48/24/6.5 см. Размери на квадратите: 4.3/4.3 см."
        self.assertEqual(scraper.parse_dimensions(text), (48.0, 24.0, 6.5))

    def test_split_weight_word_is_parsed(self):
        self.assertEqual(scraper.parse_weight("Т егло: 450 гр."), 450.0)

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

    def test_multiple_gallery_images_are_preserved(self):
        html = '''
        <div id="content"><div class="product-info"><ul class="thumbnails">
          <li><a href="https://oreshak.bg/image/catalog/a/one-1000x1000-product_popup.jpg"><img src="https://oreshak.bg/image/catalog/a/one-500x500-product_thumb.jpg"></a></li>
          <li><a href="https://oreshak.bg/image/catalog/a/two-1000x1000-product_popup.jpg"><img src="https://oreshak.bg/image/catalog/a/two-500x500-product_thumb.jpg"></a></li>
          <li><a href="https://oreshak.bg/image/catalog/a/three-1000x1000-product_popup.jpg"><img src="https://oreshak.bg/image/catalog/a/three-500x500-product_thumb.jpg"></a></li>
        </ul></div></div>
        '''
        soup = BeautifulSoup(html, "lxml")
        images = scraper.OreshakClient._parse_images(soup, {}, "https://oreshak.bg/product")
        self.assertEqual(len(images), 3)
        self.assertTrue(all("1000x1000-product_popup" in url for url in images))

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
        available, source, quantity = scraper.OreshakClient._parse_stock(soup, {})
        self.assertTrue(available)
        self.assertIsNone(quantity)
        self.assertIn("add-to-cart", source)

    def test_last_item_sets_quantity_one(self):
        html = '''
        <div id="content"><div class="product-info">
          <h1>Картина</h1><p>Последна бройка</p><button id="button-cart">Купи</button>
        </div></div>
        '''
        soup = BeautifulSoup(html, "lxml")
        available, _, quantity = scraper.OreshakClient._parse_stock(soup, {})
        self.assertTrue(available)
        self.assertEqual(quantity, 1)

    def test_plywood_does_not_trigger_textile_material(self):
        product = scraper.Product(
            url="https://oreshak.bg/p",
            source_category_url="https://oreshak.bg/komplekti-shah-i-tabla",
            title="Комплект шах и табла",
            description="Материал: липа и буков шперплат.",
            attributes={"Материал": "Бук", "Дърво": "Липа"},
        )
        candidates = scraper.detect_material_candidates(product)
        self.assertEqual(candidates[0], "Wood")
        self.assertNotEqual(candidates[0], "Textile")

    def test_dice_map_to_game_pieces_not_board_games(self):
        class FakeSchema:
            category_names = {"25613": "Game Pieces", "25615": "Board Games", "13020": "Collectible Figurines"}

        product = scraper.Product(
            url="https://oreshak.bg/dice",
            source_category_url="https://oreshak.bg/aksesoari-za-shah-i-tabla",
            source_category_name="АКСЕСОАРИ ЗА ШАХ И ТАБЛА",
            title="ЗАРЧЕТА ОТ КОСТ ЗА ТАБЛА 6.4 ММ",
            code="6348",
        )
        category_id, _, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(category_id, "25613")
        self.assertEqual(confidence, "high")

    def test_unsupported_ashtray_mapping_is_low_confidence(self):
        class FakeSchema:
            category_names = {"12140": "Collectible Figurines", "13020": "Collectible Figurines"}

        product = scraper.Product(
            url="https://oreshak.bg/ashtray",
            source_category_url="https://oreshak.bg/dyalani-unikati-ot-darvo",
            source_category_name="ДЯЛАНИ УНИКАТИ ОТ ДЪРВО",
            title="МАСИВЕН ПЕПЕЛНИК ЗА ПУРИ ОТ ОРЕХ",
        )
        _, reason, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(confidence, "low")
        self.assertIn("ashtray", reason)

    def test_acrylic_fabric_composition_is_100_percent(self):
        class FakeSchema:
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Textile Material", "Non-textile Material"] if column == "MY" else []

        product = scraper.Product(
            url="https://oreshak.bg/fabric",
            source_category_url="https://oreshak.bg/bitova-takan",
            title="БИТОВА ПОКРИВКА",
            description="Материал: 100% акрил.",
            category_id="39650",
        )
        row = scraper.fabric_composition_fields(product, FakeSchema())
        self.assertEqual(row["IA"], 100.0)
        self.assertEqual(row["MJ"], 0)
        self.assertEqual(row["MT"], 0)
        self.assertEqual(row["MY"], "Textile Material")

    def test_upload_uses_pre_promotion_price_when_enabled(self):
        product = scraper.Product(
            url="https://oreshak.bg/promo",
            source_category_url="https://oreshak.bg/komplekti-shah-i-tabla",
            title="Промо продукт",
            price_eur=Decimal("77.72"),
            list_price_eur=Decimal("81.81"),
            price_source="checkout price",
            list_price_source="visible regular price",
        )
        price, list_price, basis = scraper.prices_for_upload(product, {
            "price_multiplier": 1,
            "use_pre_promotion_price": True,
        })
        self.assertEqual(price, Decimal("81.81"))
        self.assertIsNone(list_price)
        self.assertEqual(basis, "visible regular price")

    def test_upload_uses_regular_current_price_without_promotion(self):
        product = scraper.Product(
            url="https://oreshak.bg/regular",
            source_category_url="https://oreshak.bg/komplekti-shah-i-tabla",
            title="Редовен продукт",
            price_eur=Decimal("127.82"),
            price_source="visible product price",
        )
        price, list_price, basis = scraper.prices_for_upload(product, {
            "price_multiplier": 1,
            "use_pre_promotion_price": True,
        })
        self.assertEqual(price, Decimal("127.82"))
        self.assertIsNone(list_price)
        self.assertEqual(basis, "visible product price")

    def test_discounted_price_mode_can_be_restored_from_config(self):
        product = scraper.Product(
            url="https://oreshak.bg/promo",
            source_category_url="https://oreshak.bg/komplekti-shah-i-tabla",
            title="Промо продукт",
            price_eur=Decimal("77.72"),
            list_price_eur=Decimal("81.81"),
        )
        price, list_price, _ = scraper.prices_for_upload(product, {
            "price_multiplier": 1,
            "use_pre_promotion_price": False,
        })
        self.assertEqual(price, Decimal("77.72"))
        self.assertEqual(list_price, Decimal("81.81"))

    def test_two_dimensional_size_does_not_invent_thickness(self):
        self.assertIsNone(scraper.parse_dimensions("Размери: 40/40 см. Материал: 100% акрил."))

    def test_internal_depth_is_not_used_as_outer_dimension(self):
        text = "Размери: 40/11 см. Вътрешна дълбочина: 7 см."
        self.assertIsNone(scraper.parse_dimensions(text))

    def test_external_depth_can_complete_two_dimensions(self):
        text = "Размери: 40/11 см. Дълбочина: 7 см."
        self.assertEqual(scraper.parse_dimensions(text), (40.0, 11.0, 7.0))

    def test_two_equal_tiny_dimensions_are_still_treated_as_cube(self):
        self.assertEqual(scraper.parse_dimensions("Размер: 6.4/6.4 мм."), (0.64, 0.64, 0.64))

    def test_explicit_out_of_stock_text_beats_enabled_cart_button(self):
        html = '''
        <div id="content"><div class="product-info">
          <h1>Покривка</h1><p>Няма наличност</p><button id="button-cart">Купи</button>
        </div></div>
        '''
        soup = BeautifulSoup(html, "lxml")
        available, source, quantity = scraper.OreshakClient._parse_stock(soup, {})
        self.assertFalse(available)
        self.assertEqual(quantity, 0)
        self.assertIn("availability text", source)

    def test_solid_wood_prefers_log_over_composite_wood(self):
        class FakeSchema:
            headers = {"EK": "7317 - Wood Type"}
            internal_keys = {"EK": "t_3_Wood Type"}
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Composite Wood", "Log"]
        product = scraper.Product(
            url="https://oreshak.bg/p",
            source_category_url="https://oreshak.bg/darvorezbovani-pana-i-plastiki",
            title="ПАНО ОТ ЛИПА",
            description="Материал: липа",
            attributes={"Материал": "Дърво", "Дърво": "Липа"},
            category_id="12151",
        )
        variant = scraper.Variant(product=product, option_values={}, sku="1", title="ПАНО", image="x")
        value = scraper.infer_required_value("EK", product, variant, {}, FakeSchema(), {})
        self.assertEqual(value, "Log")

    def test_product_title_beats_source_category_words(self):
        class FakeSchema:
            category_names = {"10888": "Kegs & Kegging", "10905": "Barrels", "13020": "Collectible Figurines"}

        product = scraper.Product(
            url="https://oreshak.bg/flask",
            source_category_url="https://oreshak.bg/baklitsi-i-bureta",
            source_category_name="БЪКЛИЦИ И БУРЕТА",
            title="ДЪРВОРЕЗБОВАНА БЪКЛИЦА 200 МЛ",
        )
        _, reason, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(confidence, "low")
        self.assertIn("flask", reason)

    def test_tablecloth_is_not_mapped_to_raw_fabric(self):
        class FakeSchema:
            category_names = {"39650": "Fabric", "13020": "Collectible Figurines"}

        product = scraper.Product(
            url="https://oreshak.bg/tablecloth",
            source_category_url="https://oreshak.bg/bitova-takan",
            source_category_name="БИТОВА ТЪКАН",
            title="БИТОВА ПОКРИВКА 40/40",
        )
        _, reason, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(confidence, "low")
        self.assertIn("not raw Fabric", reason)

    def test_salt_cellar_maps_to_serveware_accessory(self):
        class FakeSchema:
            category_names = {"10703": "Serveware Accessories", "9923": "Tool & Gadget Sets", "13020": "Collectible Figurines"}

        product = scraper.Product(
            url="https://oreshak.bg/salt",
            source_category_url="https://oreshak.bg/kuhnenski-aksesoari-ot-darvo-Oreshak",
            source_category_name="КУХНЕНСКИ АКСЕСОАРИ ОТ ДЪРВО",
            title="ЕДИНИЧНА ДЪРВЕНА СОЛНИЦА С КАПАК",
        )
        category_id, _, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(category_id, "10703")
        self.assertEqual(confidence, "high")

    def test_decorative_novelty_plate_is_not_assumed_food_safe(self):
        class FakeSchema:
            headers = {"FH": "4010 - Can Be Used For Food Contact"}
            internal_keys = {"FH": "t_3_Can Be Used For Food Contact"}
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Yes", "No"]

        product = scraper.Product(
            url="https://oreshak.bg/plate",
            source_category_url="https://oreshak.bg/ruchno-izraboteni-chinii-ot-darvo",
            title="ДЪРВЕНА ЧИНИЯ ПИРОГРАФИЯ С ФОЛКЛОРНИ МОТИВИ",
            description="Декоративна чиния за подарък и окачване на стена.",
            category_id="10853",
        )
        variant = scraper.Variant(product=product, option_values={}, sku="1", title=product.title, image="x")
        value = scraper.infer_required_value("FH", product, variant, {}, FakeSchema(), {})
        self.assertEqual(value, "No")


if __name__ == "__main__":
    unittest.main()
