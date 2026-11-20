from core import template


def test_every_registered_template_exists():
    specs = template.available()
    assert set(specs) == {"landscape", "portrait"}
    for spec in specs.values():
        assert spec.path.exists(), spec.path


def test_templates_satisfy_the_renderer_contract():
    for key, spec in template.available().items():
        issues = template.validate(spec.read(), key)
        assert issues == [], f"{key}: {issues}"


def test_orientation_is_reported_correctly():
    for key, spec in template.available().items():
        info = template.inspect(spec.read())
        assert info["orientation"] == key
        assert info["control_placeholder"] is True
        assert info["fonts_embedded"] is True
        assert info["tables"] >= 1
