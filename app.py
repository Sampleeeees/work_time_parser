"""Litestar base api with endpoints."""
from typing import Annotated

from litestar import post, Litestar, MediaType
from litestar.datastructures import UploadFile
from litestar.enums import RequestEncodingType
from litestar.openapi import OpenAPIConfig
from litestar.openapi.plugins import SwaggerRenderPlugin
from litestar.params import Body
from litestar.response import File

from excel_parser import ExcelParser


@post("/update-time-sheet")
async def update_timesheet(
    data: Annotated[UploadFile, Body(media_type=RequestEncodingType.MULTI_PART, title="File Upload!")],
    rate: Annotated[int, Body(title="Set your rate")],
    exchange_rate: Annotated[float, Body(title="Set your exchange rate")],
) -> File:
    """Update timesheet."""
    import tempfile

    # Save uploaded file to a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        contents = await data.read()
        tmp.write(contents)
        tmp.flush()
        temp_path = tmp.name

    # Process file
    parser = ExcelParser(workbook_path=temp_path, rate=rate, exchange_rate=exchange_rate)
    parser.remove_unused_columns()
    output_path = parser.generate_financial_report()  # should return processed file path

    return File(
        path=output_path,
        filename="financial_report.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )




app = Litestar(
    route_handlers=[update_timesheet],
    openapi_config=OpenAPIConfig(
        title="My API",
        version="1.0.0",
        path="/",
        render_plugins=[SwaggerRenderPlugin()],
    ),
)