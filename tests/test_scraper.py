import unittest
from decimal import Decimal
from bs4 import BeautifulSoup

import scraper


class ParserTests(unittest.TestCase):
    def test_project_version_matches_bundle(self):
        self.assertEqual(scraper.PROJECT_VERSION, "6.6")

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


    def test_main_price_nearest_cart_beats_unlabeled_focus_carousel(self):
        html = """
        <div id="content"><div class="product-right">
          <div class="owl-carousel focus-products">
            <ul class="list-unstyled"><li><h2 class="price">65.45 €</h2></li></ul>
          </div>
          <div class="main-product-panel">
            <h1>КОМПЛЕКТ ШАХ И ТАБЛА БУК ПИРОГРАФ 48/48</h1>
            <ul class="list-unstyled">
              <li>Код на продукта: 5076</li>
              <li><h2 class="price">81.81 €</h2></li>
            </ul>
            <button id="button-cart">Купи</button>
          </div>
        </div></div>
        """
        soup = BeautifulSoup(html, "lxml")
        price, list_price, source, _ = scraper.OreshakClient._parse_prices(soup, {})
        self.assertEqual(price, Decimal("81.81"))
        self.assertIsNone(list_price)
        self.assertIn("ul.list-unstyled h2", source)


    def test_exact_product_code_anchor_rejects_focus_price_in_same_column(self):
        html = """
        <div id="content"><div class="product-right">
          <div class="main-and-focus-wrapper">
            <div class="focus-products">
              <h2 class="price">65.45 €</h2>
            </div>
            <div class="product-buy-box">
              <h1>КОМПЛЕКТ ШАХ И ТАБЛА БУК ПИРОГРАФ 48/48</h1>
              <ul class="list-unstyled">
                <li><h2 class="price"><span>73.63 €</span> <span>81.81 €</span></h2></li>
                <li>Код на продукта: 5076</li>
              </ul>
              <button id="button-cart">Купи</button>
            </div>
          </div>
        </div></div>
        """
        soup = BeautifulSoup(html, "lxml")
        price, list_price, source, list_source = scraper.OreshakClient._parse_prices(
            soup, {}, "5076", "КОМПЛЕКТ ШАХ И ТАБЛА БУК ПИРОГРАФ 48/48"
        )
        self.assertEqual(price, Decimal("73.63"))
        self.assertEqual(list_price, Decimal("81.81"))
        self.assertIn("anchored main product selector", source)
        self.assertIn("pre-promotion", list_source)

    def test_product_code_distance_beats_unrelated_h2_when_common_parent_contains_cart(self):
        html = """
        <div id="content"><div class="product-right">
          <div><h2 class="price">65.45 €</h2></div>
          <div class="details"><h1>Main product</h1>
            <ul class="list-unstyled"><li><h2 class="price">81.81 €</h2></li><li>Код на продукта: 5076</li></ul>
          </div>
          <button id="button-cart">Купи</button>
        </div></div>
        """
        soup = BeautifulSoup(html, "lxml")
        price, list_price, _, _ = scraper.OreshakClient._parse_prices(soup, {}, "5076", "Main product")
        self.assertEqual(price, Decimal("81.81"))
        self.assertIsNone(list_price)

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

    def test_strict_measurements_reject_missing_weight_and_dimensions(self):
        product = scraper.Product(
            url="https://oreshak.bg/painting",
            source_category_url="https://oreshak.bg/kartini-ot-bulgaria",
            title="КАРТИНА",
        )
        errors = scraper.source_measurement_errors(product, {
            "omit_rows_with_fallback_measurements": True,
            "omit_rows_with_ambiguous_set_dimensions": True,
        })
        self.assertTrue(any("weight" in error.casefold() for error in errors))
        self.assertTrue(any("dimensions" in error.casefold() for error in errors))

    def test_non_strict_measurements_allow_category_fallback(self):
        product = scraper.Product(
            url="https://oreshak.bg/painting",
            source_category_url="https://oreshak.bg/kartini-ot-bulgaria",
            title="КАРТИНА",
        )
        errors = scraper.source_measurement_errors(product, {
            "omit_rows_with_fallback_measurements": False,
            "omit_rows_with_ambiguous_set_dimensions": False,
        })
        self.assertEqual(errors, [])

    def test_multi_piece_dice_dimensions_are_not_used_as_package_dimensions(self):
        product = scraper.Product(
            url="https://oreshak.bg/dice",
            source_category_url="https://oreshak.bg/aksesoari-za-shah-i-tabla",
            title="ЗАРЧЕТА ОТ КОСТ ЗА ТАБЛА 6.4 ММ",
            description="Един комплект съдържа 2 броя зарчета с размер 6.4 мм. на страна.",
            weight_g=1.0,
            dimensions_cm=(0.64, 0.64, 0.64),
        )
        errors = scraper.source_measurement_errors(product, {
            "omit_rows_with_fallback_measurements": True,
            "omit_rows_with_ambiguous_set_dimensions": True,
        })
        self.assertTrue(any("2 pieces" in error for error in errors))


    def test_estimated_package_measurements_add_padding_and_review_notes(self):
        product = scraper.Product(
            url="https://oreshak.bg/item",
            source_category_url="https://oreshak.bg/aksesoari-za-lovni-trofei",
            title="ДЪРВЕН ПРОДУКТ",
            category_id="32650",
            weight_g=1000.0,
            dimensions_cm=(30.0, 20.0, 10.0),
        )
        weight, dims, notes, weight_basis, dims_basis = scraper.package_measurements_for_upload(product, {
            "package_measurement_mode": "estimate_and_review",
            "package_dimension_padding_cm": 2.0,
            "package_weight_padding_percent": 10.0,
            "package_weight_padding_min_g": 50.0,
        })
        self.assertEqual(weight, 1100.0)
        self.assertEqual(dims, (32.0, 22.0, 12.0))
        self.assertTrue(notes)
        self.assertIn("estimated package weight", weight_basis)
        self.assertIn("estimated package dimensions", dims_basis)

    def test_estimated_package_measurements_use_category_fallback_when_missing(self):
        product = scraper.Product(
            url="https://oreshak.bg/item",
            source_category_url="https://oreshak.bg/aksesoari-za-lovni-trofei",
            title="ДЪРВЕН ПРОДУКТ",
            category_id="32650",
        )
        weight, dims, notes, weight_basis, dims_basis = scraper.package_measurements_for_upload(product, {
            "package_measurement_mode": "estimate_and_review",
        })
        self.assertEqual(weight, 600.0)
        self.assertEqual(dims, (35.0, 25.0, 5.0))
        self.assertTrue(notes)
        self.assertIn("category fallback", weight_basis)
        self.assertIn("category fallback", dims_basis)
        self.assertEqual(scraper.source_measurement_errors(product, {"package_measurement_mode": "estimate_and_review"}), [])

    def test_ambiguous_set_dimensions_use_fallback_and_review(self):
        product = scraper.Product(
            url="https://oreshak.bg/dice",
            source_category_url="https://oreshak.bg/aksesoari-za-shah-i-tabla",
            title="ЗАРЧЕТА ОТ КОСТ ЗА ТАБЛА 6.4 ММ",
            description="Един комплект съдържа 2 броя зарчета с размер 6.4 мм. на страна.",
            category_id="25613",
            weight_g=1.0,
            dimensions_cm=(0.64, 0.64, 0.64),
        )
        weight, dims, notes, _, dims_basis = scraper.package_measurements_for_upload(product, {
            "package_measurement_mode": "estimate_and_review",
        })
        self.assertEqual(weight, 51.0)
        self.assertEqual(dims, (20.0, 15.0, 8.0))
        self.assertTrue(notes)
        self.assertIn("2 pieces", dims_basis)

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


    def test_regulated_consumables_are_rejected(self):
        class FakeSchema:
            category_names = {"13020": "Collectible Figurines"}
        for title in (
            "ГЮЛОВА РАКИЯ 100 МЛ",
            "ЛИКЬОР ОТ РОЗИ 100 МЛ",
            "НАТУРАЛНО СЛАДКО ОТ РОЗИ 20 МЛ",
            "БИО НАТУРАЛНА РОЗОВА ВОДА 100 МЛ",
            "НАТУРАЛНО ЛАВАНДУЛОВО МАСЛО 10 МЛ",
            "СУВЕНИРНА КАРТИЧКА С ПАРФЮМ БЪЛГАРСКА РОЗА",
            "КРЕМ ЗА РЪЦЕ РОЗА 75 МЛ",
        ):
            product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/suveniri", title=title)
            _, _, confidence = scraper.category_for(product, {}, FakeSchema())
            self.assertEqual(confidence, "low", title)

    def test_smoking_accessories_are_rejected(self):
        class FakeSchema:
            category_names = {"12179": "Decorative Boxes"}
        for title in ("ХУМИДОР ЗА ПУРИ С ВЛАГОМЕР", "ПОДАРЪЧЕН КОМПЛЕКТ БЕНЗИНОВА ЗАПАЛКА"):
            product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/kutii-za-aksesoari", title=title)
            _, _, confidence = scraper.category_for(product, {}, FakeSchema())
            self.assertEqual(confidence, "low", title)

    def test_belts_sheaths_and_jewelry_are_rejected(self):
        class FakeSchema:
            category_names = {"10059": "Chef Knives", "13020": "Collectible Figurines"}
        cases = [
            ("МЪЖКИ КОЛАН ОТ КОЖА", "https://oreshak.bg/nojove-ot-balgaria"),
            ("КОЖЕНА КАНИЯ ЗА НОЖ 30 СМ", "https://oreshak.bg/nojove-ot-balgaria"),
            ("РЪЧНО ИЗРАБОТЕНО КОЛИЕ ОТ МЪНИСТА", "https://oreshak.bg/suveniri"),
            ("РЪЧНО ИЗРАБОТЕНА ГРИВНА ОТ МЪНИСТА", "https://oreshak.bg/suveniri"),
        ]
        for title, url in cases:
            product = scraper.Product(url="https://oreshak.bg/x", source_category_url=url, title=title)
            _, _, confidence = scraper.category_for(product, {}, FakeSchema())
            self.assertEqual(confidence, "low", title)

    def test_functional_home_items_are_not_figurines(self):
        class FakeSchema:
            category_names = {"12140": "Collectible Figurines"}
        for title in ("ДЪРВЕН СТЕНЕН ЧАСОВНИК УНИКАТ", "СТОЙКА ЗА КЛЮЧОВЕ", "ДЪРВЕНА АРТ ЗАКАЧАЛКА"):
            product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/dyalani-unikati-ot-darvo", title=title)
            _, _, confidence = scraper.category_for(product, {}, FakeSchema())
            self.assertEqual(confidence, "low", title)

    def test_unsupported_kitchen_items_are_rejected(self):
        class FakeSchema:
            category_names = {"9923": "Tool & Gadget Sets"}
        for title in ("ДЪРВЕНА ТОЧИЛКА", "ХАВАНЧЕ ЗА ПОДПРАВКИ", "ДЪРВЕНА ХАЛБА", "СТОМАНЕН ШИШ", "ДЪРВЕНА КУПА"):
            product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/kuhnenski-aksesoari-ot-darvo-Oreshak", title=title)
            _, _, confidence = scraper.category_for(product, {}, FakeSchema())
            self.assertEqual(confidence, "low", title)

    def test_kitchen_spoon_and_board_map_explicitly(self):
        class FakeSchema:
            category_names = {"9999": "Cooking Spoons", "54423": "Serving Boards", "9923": "Tool & Gadget Sets"}
        spoon = scraper.Product(url="https://oreshak.bg/s", source_category_url="https://oreshak.bg/kuhnenski-aksesoari-ot-darvo-Oreshak", title="ДЪРВЕНА ЛЪЖИЧКА ЗА МЕД")
        board = scraper.Product(url="https://oreshak.bg/b", source_category_url="https://oreshak.bg/kuhnenski-aksesoari-ot-darvo-Oreshak", title="ПРЕМИУМ КУХНЕНСКА ДЪСКА ОТ БУК")
        self.assertEqual(scraper.category_for(spoon, {}, FakeSchema())[0:3:2], ("9999", "high"))
        self.assertEqual(scraper.category_for(board, {}, FakeSchema())[0:3:2], ("54423", "high"))

    def test_single_wine_stopper_is_rejected(self):
        class FakeSchema:
            category_names = {"39880": "Gift Boxes", "10875": "Wine Accessory Sets"}
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/kutii-za-vino-i-bijuta", title="ЗАПУШАЛКА ЗА БУТИЛКА ВИНО ОТ ЕЛЕНОВ РОГ")
        _, reason, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(confidence, "low")
        self.assertIn("stopper", reason)

    def test_unknown_product_in_heterogeneous_source_requires_explicit_rule(self):
        class FakeSchema:
            category_names = {"13020": "Collectible Figurines"}
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/suveniri", title="НЕОБИЧАЕН ПРЕДМЕТ")
        _, reason, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(confidence, "low")
        self.assertIn("explicit safe product rule", reason)

    def test_real_household_knife_maps_high_confidence(self):
        class FakeSchema:
            category_names = {"10059": "Chef Knives", "10072": "Utility Knives"}
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/nojove-ot-balgaria", title="ДОМАКИНСКИ НОЖ ЗА СИРЕНА")
        category_id, _, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(category_id, "10059")
        self.assertEqual(confidence, "high")

    def test_knife_surface_finish_defaults_to_polishing(self):
        class FakeSchema:
            headers = {"GT": "1081 - Surface Finishing Type"}
            internal_keys = {"GT": "t_3_Surface Finishing Type"}
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Matte", "Polishing", "Hammered"]
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/nojove-ot-balgaria", title="ДОМАКИНСКИ НОЖ", category_id="10059")
        variant = scraper.Variant(product=product, option_values={}, sku="1", title=product.title, image="x")
        self.assertEqual(scraper.infer_required_value("GT", product, variant, {}, FakeSchema(), {}), "Polishing")
        self.assertTrue(any("Surface finishing" in warning for warning in product.warnings))

    def test_forged_knife_surface_finish_maps_to_hammered(self):
        class FakeSchema:
            headers = {"GT": "1081 - Surface Finishing Type"}
            internal_keys = {"GT": "t_3_Surface Finishing Type"}
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Matte", "Polishing", "Hammered"]
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/nojove-ot-balgaria", title="РЪЧНО КОВАН ЛОВЕН НОЖ", category_id="10072")
        variant = scraper.Variant(product=product, option_values={}, sku="1", title=product.title, image="x")
        self.assertEqual(scraper.infer_required_value("GT", product, variant, {}, FakeSchema(), {}), "Hammered")

    def test_horn_handle_uses_traceable_review_fallback(self):
        class FakeSchema:
            headers = {"LV": "7307 - Handle Material"}
            internal_keys = {"LV": "t_3_Handle Material"}
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Metal", "Wooden handle", "Resin"]
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/nojove-ot-balgaria", title="ЛОВЕН НОЖ С ДРЪЖКА ОТ ЕЛЕНОВ РОГ", category_id="10072")
        variant = scraper.Variant(product=product, option_values={}, sku="1", title=product.title, image="x")
        self.assertEqual(scraper.infer_required_value("LV", product, variant, {}, FakeSchema(), {}), "Resin")
        self.assertTrue(any("horn/bone" in warning for warning in product.warnings))

    def test_missing_description_is_generated_from_published_data(self):
        product = scraper.Product(
            url="https://oreshak.bg/x",
            source_category_url="https://oreshak.bg/kutii-za-aksesoari",
            title="ДЪРВЕНА КУТИЯ ЗА БИЖУТА",
            code="6360",
            attributes={"Материал": "Дърво", "Размер": "20 x 15 x 10 см"},
        )
        self.assertTrue(scraper.ensure_product_description(product))
        self.assertIn("ДЪРВЕНА КУТИЯ ЗА БИЖУТА", product.description)
        self.assertIn("Материал: Дърво", product.description)
        self.assertTrue(product.warnings)


    def test_material_word_does_not_trigger_matte_finish(self):
        class FakeSchema:
            headers = {"GT": "1081 - Surface Finishing Type"}
            internal_keys = {"GT": "t_3_Surface Finishing Type"}
            @staticmethod
            def dropdown_for(column, category_id, row):
                return ["Matte", "Polishing", "Hammered"]
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/nojove-ot-balgaria", title="ДОМАКИНСКИ НОЖ", description="Материал: стомана.", category_id="10059")
        variant = scraper.Variant(product=product, option_values={}, sku="1", title=product.title, image="x")
        self.assertEqual(scraper.infer_required_value("GT", product, variant, {}, FakeSchema(), {}), "Polishing")

    def test_white_blank_board_keeps_art_board_mapping(self):
        class FakeSchema:
            category_names = {"39981": "Wood Art Boards", "12193": "Decorative Plaques"}
        product = scraper.Product(url="https://oreshak.bg/x", source_category_url="https://oreshak.bg/Wooden-souvenirs-white-blank", title="СУВЕНИРНА ДЪСКА ЗА ДЕКУПАЖ НА БЯЛА ЗАГОТОВКА")
        category_id, _, confidence = scraper.category_for(product, {}, FakeSchema())
        self.assertEqual(category_id, "39981")
        self.assertEqual(confidence, "high")


    def test_wooden_source_category_supports_board_material_when_page_omits_it(self):
        product = scraper.Product(
            url="https://oreshak.bg/kuhnenski-aksesoari-ot-darvo-Oreshak/board",
            source_category_url="https://oreshak.bg/kuhnenski-aksesoari-ot-darvo-Oreshak",
            source_category_name="КУХНЕНСКИ АКСЕСОАРИ ОТ ДЪРВО",
            title="ПРОФЕСИОНАЛНА ДЪСКА ЗА РЯЗАНЕ НА ХЛЯБ",
            description="Размери: 47.5/23/2.5 см.",
        )
        self.assertEqual(scraper.detect_material_candidates(product)[0], "Wood")


if __name__ == "__main__":
    unittest.main()
