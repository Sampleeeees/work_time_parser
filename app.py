"""Litestar base api with endpoints."""
import tempfile
from typing import Annotated

from litestar import post, Litestar
from litestar.datastructures import UploadFile
from litestar.enums import RequestEncodingType
from litestar.openapi import OpenAPIConfig
from litestar.openapi.plugins import SwaggerRenderPlugin
from litestar.params import Body
from litestar.response import File

from enums import ReportGroupingType
from excel_parser import ExcelParser


@post("/generate-personal-time")
async def generate_personal_time(
    data: Annotated[UploadFile, Body(media_type=RequestEncodingType.MULTI_PART, title="File Upload!")],
    rate: Annotated[int, Body(title="Set your rate")],
    exchange_rate: Annotated[float, Body(title="Set your exchange rate")]
) -> File:
    """Update timesheet."""

    # Save uploaded file to a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        contents = await data.read()
        tmp.write(contents)
        tmp.flush()
        temp_path = tmp.name

    # Process file
    parser = ExcelParser(workbook_path=temp_path)
    output_path = parser.generate_financial_report(rate=rate, exchange_rate=exchange_rate)

    return File(
        path=output_path,
        filename=output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@post("/generate-project-time")
async def generate_project_time(
    data: Annotated[UploadFile, Body(media_type=RequestEncodingType.MULTI_PART, title="File Upload!")],
    group_type: ReportGroupingType
) -> File:
    """Generate project time."""
    # Save uploaded input file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        contents = await data.read()
        tmp_input.write(contents)
        tmp_input.flush()
        input_path = tmp_input.name

    # Prepare output file path
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
        output_path = tmp_output.name

    # Process Excel
    parser = ExcelParser(workbook_path=input_path)
    parser.generate_project_report(group_type=group_type)

    # Return and delete after response
    return File(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path
    )


app = Litestar(
    route_handlers=[generate_personal_time, generate_project_time],
    openapi_config=OpenAPIConfig(
        title="My API",
        version="1.0.0",
        path="/",
        render_plugins=[SwaggerRenderPlugin()],
    ),
    debug=True
)