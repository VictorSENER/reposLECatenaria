/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.controller;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.sener.sireca.web.service.AuthenticationService;
import com.sener.sireca.web.session.UserCredential;

@Controller
public class AuthenticationController
{
    @Autowired
    AuthenticationService authService;

    @RequestMapping(value = { "", "/", "index" }, method = RequestMethod.GET)
    public String index(ModelMap model, HttpServletRequest request,
            HttpSession session)
    {
        UserCredential cre = authService.getUserCredential(session);
        if (cre != null)
            return "redirect:main";
        else
            return "redirect:login";
    }

    @RequestMapping(value = "login", method = RequestMethod.GET)
    public String login(Model model)
    {
        return "login.zul";
    }

    @RequestMapping(value = "logout", method = RequestMethod.GET)
    public String logout(ModelMap model, HttpServletRequest request,
            HttpSession session)
    {
        authService.logout(session);
        return "redirect:index";
    }

    @RequestMapping(value = "main", method = RequestMethod.GET)
    public String main(Model model)
    {
        return "main.zul";
    }

    @RequestMapping(value = "user", method = RequestMethod.GET)
    public String user(Model model)
    {
        return "user.zul";
    }

    @RequestMapping(value = "catenary", method = RequestMethod.GET)
    public String catenary(Model model)
    {
        return "catenary.zul";
    }

    @RequestMapping(value = "project", method = RequestMethod.GET)
    public String project(Model model)
    {
        return "project.zul";
    }

    @RequestMapping(value = "project/edit/{id}", method = RequestMethod.GET)
    public String editProject(@PathVariable Integer id, Model model)
    {
        return "projectEdit.zul";
    }

    @RequestMapping(value = "replanteo", method = RequestMethod.GET)
    public String replanteo(Model model)
    {
        return "replanteo.zul";
    }

    @RequestMapping(value = "drawing", method = RequestMethod.GET)
    public String drawing(Model model)
    {
        return "drawing.zul";
    }

    @RequestMapping(value = "pendolado", method = RequestMethod.GET)
    public String pendolado(Model model)
    {
        return "pendolado.zul";
    }

    @RequestMapping(value = "montaje", method = RequestMethod.GET)
    public String montaje(Model model)
    {
        return "montaje.zul";
    }
}
