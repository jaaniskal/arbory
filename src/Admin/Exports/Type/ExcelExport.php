<?php

namespace Arbory\Base\Admin\Exports\Type;

use Arbory\Base\Admin\Exports\DataSetExport;
use Arbory\Base\Admin\Exports\ExportInterface;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

class ExcelExport implements ExportInterface
{
    const EXTENSION = 'xlsx';

    public function __construct(
        protected DataSetExport $export
    ) {
    }

    public function download(string $fileName): BinaryFileResponse
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->fromArray($this->export->getColumns(), null, 'A1');
        $sheet->fromArray($this->export->getItems()->toArray(), null, 'A2');

        $tempPath = tempnam(sys_get_temp_dir(), 'arbory_export_');

        (new Xlsx($spreadsheet))->save($tempPath);
        $spreadsheet->disconnectWorksheets();

        return response()->download($tempPath, $fileName . '.' . self::EXTENSION)->deleteFileAfterSend();
    }
}
