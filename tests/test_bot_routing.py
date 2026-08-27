import bot


def test_services_are_grouped_into_categories():
    assert len(bot.SERVICES) == 18
    category_ids = {category["id"] for category in bot.SERVICE_CATEGORIES}
    assert category_ids == {"platforms", "documents", "images", "data", "tools"}
    assert all(service["category"] in category_ids for service in bot.SERVICES)
    assert len(bot.services_for_category("platforms")) == 2
    assert len(bot.services_for_category("documents")) == 7
    assert len(bot.services_for_category("images")) == 3
    assert len(bot.services_for_category("data")) == 5
    assert len(bot.services_for_category("tools")) == 1


def test_number_selection_is_local_to_the_chosen_category():
    assert bot.service_by_category_number("documents", 1)["id"] == "format_word"
    assert bot.service_by_category_number("documents", 2)["id"] == "merge_pdf"
    assert bot.service_by_category_number("data", 1)["id"] == "excel_to_vcf"
    assert bot.service_by_category_number("data", 2)["id"] == "clean_excel"
    assert bot.service_by_category_number("data", 4)["id"] == "merge_excel"
    assert bot.service_by_category_number("data", 5)["id"] == "colornote_to_html"
    assert bot.service_by_category_number("data", 6) is None


def test_initial_screen_contains_categories_not_services():
    initial_text = bot.categories_text()
    assert "اختر الصنف" in initial_text
    assert "تحويل Excel إلى VCF" not in initial_text
    callbacks = [
        button.callback_data
        for row in bot.build_categories_keyboard().inline_keyboard
        for button in row
    ]
    assert callbacks == [
        "cat:platforms",
        "cat:documents",
        "cat:images",
        "cat:data",
        "cat:tools",
    ]


def test_processing_limits_are_valid():
    assert bot.MAX_MULTI_FILES >= 2
