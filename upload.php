<?php
// 提交个人信息页面，个人基本信息（姓名、性别、专业、学院、qq号、手机号码）必填、部门申请（一志愿）必填，备注不必须填写）
// 数据库内容只包括提交个人基本信息、备注部分，
// 如果一个人填写多次报名表，以最后一次为准（这里以电话号码为主键）
// 根据同学的志愿结果，将他的报名表发送到对应部门
// 包含数据库连接文件
include 'db_connect.php'; // 连接到数据库


function submitApplicantInfo($postData) {
    // 验证必填字段
    $requiredFields = ['name', 'sex', 'major', 'college', 'QQ', 'number', 'first_choice'];
    foreach ($requiredFields as $field) {
        if (empty($postData[$field])) {
            header('Content-Type: application/json');
            return json_encode(['error' => "缺少必填字段：{$field},111"]);
        }
    }

    // 捕获表单数据
    $name = $postData['name'];
    $sex = $postData['sex'];
    $major = $postData['major'];
    $college = $postData['college'];
    $QQ = $postData['QQ'];
    $number = $postData['number'];
    $notes = isset($postData['notes']) ? $postData['notes'] : '';
    $first_choice = $postData['first_choice'];
    $second_choice = isset($postData['second_choice']) ? $postData['second_choice'] : '';
    $third_choice = isset($postData['third_choice']) ? $postData['third_choice'] : '';

    // 使用 mysqli 进行数据库操作
    global $conn;

    // 准备 SQL 语句
    $sql = "REPLACE INTO newstudent (name, sex, major, college, QQ, number, notes, first_choice) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
    $stmt = $conn->prepare($sql);

    // 绑定参数并执行
    $stmt->bind_param('ssssssss', $name, $sex, $major, $college, $QQ, $number, $notes, $first_choice);
    $stmt->execute();
    $stmt->close();

    // 获取部门信息
    $department = [
        'first_choice' => $first_choice,
        'second_choice' => $second_choice,
        'third_choice' => $third_choice,
    ];

    // 发送报名表到对应部门
    sendToDepartment($first_choice, $postData);
    if ($second_choice) {
        sendToDepartment($second_choice, $postData);
    }
    if ($third_choice) {
        sendToDepartment($third_choice, $postData);
    }

    // JSON化部门信息
    header('Content-Type: application/json');
    return json_encode($department);
}


require 'D:/phpstudy_pro/WWW/www.newstudent.com/newstudent/vendor/autoload.php'; // 确保加载 Composer 的自动加载文件
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function sendToDepartment($department, $postData) {
    // 获取部门的存储路径
    $departmentPath = getDepartmentPath($department);

    if (!$departmentPath) {
        echo json_encode(['error' => '无效的部门,222']);
        return;
    }

    // 创建新的 Spreadsheet 对象
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // 设置表头
    $sheet->setCellValue('A1', '姓名');
    $sheet->setCellValue('B1', '性别');
    $sheet->setCellValue('C1', '专业');
    $sheet->setCellValue('D1', '学院');
    $sheet->setCellValue('E1', 'QQ');
    $sheet->setCellValue('F1', '电话');
    $sheet->setCellValue('G1', '备注');
    $sheet->setCellValue('H1', '第一志愿');
    $sheet->setCellValue('I1', '第二志愿');
    $sheet->setCellValue('J1', '第三志愿');

    // 填充数据
    $sheet->setCellValue('A2', $postData['name']);
    $sheet->setCellValue('B2', $postData['sex']);
    $sheet->setCellValue('C2', $postData['major']);
    $sheet->setCellValue('D2', $postData['college']);
    $sheet->setCellValue('E2', $postData['QQ']);
    $sheet->setCellValue('F2', $postData['number']);
    $sheet->setCellValue('G2', $postData['notes']);
    $sheet->setCellValue('H2', $postData['first_choice']);
    $sheet->setCellValue('I2', $postData['second_choice']);
    $sheet->setCellValue('J2', $postData['third_choice']);

    // 保存文件
    $writer = new Xlsx($spreadsheet);
    $filePath = $departmentPath . '.xlsx';
    $writer->save($filePath);
}


function getDepartmentPath($department) {
    // 这里应该根据部门名获取对应的存储路径
    $departmentPaths = [
        '产品经理与产品运营部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/产品经理与产品运营部',
        '设计部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/设计部',
        '技术研发部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/技术研发部',
        '音视频文化部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/音视频文化部',
        '新闻通讯部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/新闻通讯部',
        '外宣部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/外宣部',
        '行政事务部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/行政事务部',
        '企划公关部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/企划公关部',
        '微信推文部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/微信推文部',
        '媒体运营部' => 'D:/phpstudy_pro/WWW/www.newstudent.com/媒体运营部'
    ];

    return isset($departmentPaths[$department]) ? $departmentPaths[$department] : null;
}

// 使用示例
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // 处理 POST 请求
    $response = submitApplicantInfo($_POST);

    // 返回 JSON 响应
    echo $response;
} else {
    echo json_encode(['error' => '无效的请求,333']);
}


