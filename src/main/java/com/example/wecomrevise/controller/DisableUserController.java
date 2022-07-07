package com.example.wecomrevise.controller;

import com.example.wecomrevise.service.DisableUserService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.PostConstruct;
import java.io.IOException;

@RestController
@RequiredArgsConstructor
public class DisableUserController {

    private final DisableUserService disableUserService;

    @PostConstruct
    public void getTokenController() throws IOException {
        disableUserService.startDisableUserService();
    }
}