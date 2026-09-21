package com.example.demo2.controller;

import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.nio.file.Path;
import java.nio.file.Paths;

@RestController
@RequestMapping("/api/test")
public class TestController {

    private static final String JSON_DIR = "D:/project/idea_workspace/demo2/json/";

    @GetMapping(value = "/json", produces = MediaType.APPLICATION_JSON_VALUE)
    public Resource getJson() {
        Path path = Paths.get(JSON_DIR, "data.json");
        return new FileSystemResource(path);
    }
}