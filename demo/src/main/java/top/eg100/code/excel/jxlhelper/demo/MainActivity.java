package top.eg100.code.excel.jxlhelper.demo;

import android.app.Activity;
import android.content.res.AssetManager;
import android.os.Bundle;
import android.os.Environment;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import top.eg100.code.excel.jxlhelper.ExcelManager;

public class MainActivity extends Activity {

    private Button mExport;
    private Button mImport;
    private TextView mExportResult;
    private TextView mImportResult;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();
    }

    @SuppressWarnings("unchecked")
    private <T> T $(int id) {
        return (T) findViewById(id);
    }

    private void initView() {
        mExport = $(R.id.btn_export);
        mImport = $(R.id.btn_import);
        mExportResult = $(R.id.exportresult);
        mImportResult = $(R.id.importresult);
        mExport.setOnClickListener(new View.OnClickListener() {

            @Override
            public void onClick(View v) {
                onExport();
            }
        });
        mImport.setOnClickListener(new View.OnClickListener() {

            @Override
            public void onClick(View v) {
                onImport();
            }
        });
    }

    private void onImport() {
        try {
            AssetManager asset = getAssets();
            long t1 = System.currentTimeMillis();
            InputStream excelStream = asset.open("users.xls");
            excelStream = new FileInputStream("users.xls");
            ExcelManager excelManager = new ExcelManager();
            List<UserExcelBean> users = excelManager.fromExcel(excelStream, UserExcelBean.class);
            long t2 = System.currentTimeMillis();
            double time = (t2 - t1) / 1000.0D;
            mImportResult.setText("读到User个数:" + users.size() + "\n用时:" + time + "秒");
        } catch (Exception e) {
            mImportResult.setText("读取异常");
            e.printStackTrace();
        }

    }

    @SuppressWarnings("unused")
    private void onExport() {
        //实际使用的时候，不要在主线程做操作，demo的数据比较少
        try {
            long t1 = System.currentTimeMillis();
            List<UserExcelBean> users = new ArrayList<>();
            for (int i = 1; i <= 150; i++) {
                UserExcelBean u = new UserExcelBean();
                u.setName("大到飞起来" + i);
                u.setMobile("手机号" + i);
                u.setSex("男");
                u.setAddress("地点" + i);
                u.setMemo("备注" + i);
                u.setOther("其他信息" + i);
                users.add(u);
            }
            String sdPath = Environment.getExternalStorageDirectory().toString();
            String filePath = sdPath + "/excel.demo/export";
            File dir = new File(filePath);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            String usersFilePath = filePath + "/users.xls";
            ExcelManager excelManager = new ExcelManager();
            OutputStream excelStream = new FileOutputStream(usersFilePath);

            boolean success = excelManager.toExcel(excelStream, users);
            long t2 = System.currentTimeMillis();

            double time = (t2 - t1) / 1000.0D;
            if (success) {
                mExportResult.setText("导出成功：在存储卡根目录:\nexcel.demo/export/users.xls" + "\n用时:" + time + "秒");
            } else {
                mExportResult.setText("导出失败");
            }

        } catch (Exception e) {
            mExportResult.setText("导出异常");
            e.printStackTrace();
        }
    }

    void ads() {
        try {
            long t1 = System.currentTimeMillis();
            List<UserExcelBean> users = new ArrayList<>();
            for (int i = 1; i <= 150; i++) {
                UserExcelBean u = new UserExcelBean();
                u.setName("大到飞起来" + i);
                u.setMobile("手机号" + i);
                u.setSex("男");
                u.setAddress("地点" + i);
                u.setMemo("备注" + i);
                u.setOther("其他信息" + i);
                users.add(u);
            }
            ExcelManager excelManager = new ExcelManager();
            OutputStream excelStream = new FileOutputStream("usersExport.xls");

            boolean success = excelManager.toExcel(excelStream, users);
            long t2 = System.currentTimeMillis();
            double time = (t2 - t1) / 1000.0D;
            if (success) {
                System.out.println("导出成功\n用时:" + time + "秒");
            } else {
                System.out.println("导出失败");
            }

        } catch (Exception e) {
            System.err.println("导出异常");
            e.printStackTrace();
        }
    }

}
