import io

from openpyxl import Workbook

from kns_plan.intake import (guess_mapping, parse_number, parse_size, read_table, stores_by_month, to_competitors,
                             to_skus, write_template)
from kns_plan.schema import Settings


def _xlsx(rows) -> bytes:
    wb = Workbook()
    for r in rows:
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_parse_number_handles_local_formats():
    assert parse_number("$9.50") == 9.5
    assert parse_number("Rp 1.186.000") == 1186000
    assert parse_number("Rp 720.000", idr=True) == 720000 and parse_number("0.125") == 0.125
    assert parse_number("1,186,000") == 1186000
    assert parse_number("9,50") == 9.5
    assert parse_number("1.234,50") == 1234.5
    assert parse_number("1,234.50") == 1234.5
    assert parse_number("") is None and parse_number("n/a") is None


def test_parse_size():
    assert parse_size("15ml") == (15, "ml", None)
    assert parse_size("31 x 4 g") == (124, "g", "31 x 4")
    assert parse_size(50) == (50, None, None)
    assert parse_size("about 15") == (None, None, None)


def test_messy_brand_file_maps_and_converts():
    data = _xlsx([
        ["Brand A price list 2027"],  # title row above the header is skipped
        [],
        ["No", "Item Code", "Product Description", "Volume", "FOB Price (USD)", "Shelf life", "Barcode"],
        [1, "PX-15", "Eye serum stick", "15ml", "$9.50", "36 months", "8809"],
        [2, "PH-60", "Collagen eye patch", "60 pcs", 7.1, "24 months", None],
        [3, "", "", "", "", "", ""],
        [4, "XX", None, "30ml", 5, None, None],
    ])
    table = read_table(data, "brand.xlsx")
    m = guess_mapping("products", table.headers)
    assert m["id"] == "Item Code" and m["name"] == "Product Description"
    assert m["size"] == "Volume" and m["fob"] == "FOB Price (USD)"
    result = to_skus(table, m)
    assert [s.name for s in result.items] == ["Eye serum stick", "Collagen eye patch"]
    first = result.items[0]
    assert (first.size, first.unit, first.fob) == (15, "ml", 9.5)
    assert first.attributes == {"No": "1", "Shelf life": "36 months", "Barcode": "8809", "Brand code": "PX-15"}
    assert len(result.skipped) == 1 and "no product name" in result.skipped[0]


def test_csv_competitors_match_our_products_by_name(full_plan):
    csv = ("Our product;Brand;Product;COO;Size;Price\n"
           "Serum;Clinique;All About Eyes;US;15 ml;Rp 720.000\n"
           "Unknown;X;Y;;;100000\n").encode()
    table = read_table(csv, "comp.csv")
    m = guess_mapping("competitors", table.headers)
    result = to_competitors(table, m, full_plan)
    assert len(result.items) == 1 and result.items[0].sku == "serum" and result.items[0].price_idr == 720000
    assert result.items[0].size == 15 and result.items[0].unit == "ml"
    assert "does not match" in result.skipped[0]


def test_store_list_becomes_monthly_counts():
    data = _xlsx([["Channel", "Store", "Opening month"],
                  ["Watsons", "Grand Indonesia", "2027-01"], ["Watsons", "PIM", "Mar 2027"],
                  ["Sephora", "PS", 2], ["Watsons", "Late", "2030-01"]])
    table = read_table(data, "stores.xlsx")
    m = guess_mapping("stores", table.headers)
    counts = stores_by_month(table, m, Settings(brand="x", fx_rate=1, start_year=2027, start_month=1, years=1))
    assert counts["Watsons"][:4] == [1, 1, 2, 2] and counts["Watsons"][-1] == 2
    assert counts["Sephora"][:3] == [0, 1, 1]


def test_template_round_trips_through_the_mapper():
    table = read_table(write_template(), "template.xlsx")
    m = guess_mapping("products", table.headers)
    assert all(m[f] for f in ("id", "name", "size", "unit", "fob", "target_retail", "moq"))
