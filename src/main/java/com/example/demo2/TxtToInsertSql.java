package com.example.demo2;

import java.io.*;
import java.nio.file.*;

public class TxtToInsertSql {

    public static void main(String[] args) {
        String folderPath = "指定文件夹路径"; // 改成你的文件夹路径
        String outputFile = folderPath + "/output.sql";

        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".txt"));

        if (files == null || files.length == 0) {
            System.out.println("没有找到txt文件");
            return;
        }

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile))) {

            for (File file : files) {
                try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
                    String line;
                    boolean firstLine = true;

                    while ((line = reader.readLine()) != null) {
                        line = line.trim();
                        if (line.isEmpty()) continue;

                        // 第一行跳过
                        if (firstLine) {
                            firstLine = false;
                            continue;
                        }

                        // 拼接SQL并写入
                        String sql = "INSERT INTO `trd_rpt`.`trd_rpt`(`id`, `name`, `lock`) VALUES (" + line + ");";
                        writer.write(sql);
                        writer.newLine();
                    }

                } catch (IOException e) {
                    System.out.println("读取失败: " + file.getName() + " - " + e.getMessage());
                }
            }

            System.out.println("SQL已生成到: " + outputFile);

        } catch (IOException e) {
            System.out.println("写入失败: " + e.getMessage());
        }
    }
}