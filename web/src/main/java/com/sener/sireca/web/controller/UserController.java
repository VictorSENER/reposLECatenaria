/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.sener.sireca.web.bean.User;
import com.sener.sireca.web.service.UserService;

@Controller
public class UserController
{

    @Autowired
    UserService userService;

    @RequestMapping(value = "user/list", method = RequestMethod.GET)
    public String listUsers(Model model)
    {
        List<User> allUsers = userService.getAllUsers();
        model.addAttribute("userList", allUsers);
        return "user.zul";
    }

    @RequestMapping(value = "user/new", method = RequestMethod.GET)
    public String newUser(Model model)
    {
        model.addAttribute("userObj", new User());
        return "form.jsp";
    }

    @RequestMapping(value = "user/edit/{id}", method = RequestMethod.GET)
    public String editUser(@PathVariable Integer id, Model model)
    {
        User userObj = userService.getUserById(id);
        model.addAttribute("userObj", userObj);
        return "form.jsp";
    }

    @RequestMapping(value = "user/save", method = RequestMethod.POST)
    public String saveUser(@ModelAttribute("userObj") User userObj,
            BindingResult br, Model model)
    {
        if (br.hasErrors())
            return "form";

        if (userObj.getId() == null)
            userService.updateUser(userObj);
        else
            userService.insertUser(userObj);

        return "redirect:/user/list";
    }

    @RequestMapping(value = "user/delete/{id}", method = RequestMethod.GET)
    public String deleteUser(@PathVariable Integer id)
    {
        userService.deleteUser(id);
        return "redirect:/user/list";
    }
}
