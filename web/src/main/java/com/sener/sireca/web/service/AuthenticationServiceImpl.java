/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.User;
import com.sener.sireca.web.session.UserCredential;

@Service("authService")
public class AuthenticationServiceImpl implements AuthenticationService
{

    @Autowired
    UserService userService;

    public UserCredential getUserCredential(HttpSession session)
    {
        // Get user credentials from current session.
        UserCredential cre = (UserCredential) session.getAttribute("userCredential");
        if (cre == null)
            return null;

        // Credentials found.
        // Check if they correspond to a valid user.
        User user = userService.getUserByUsername(cre.getUsername());
        if (user == null)
            return null;

        // Check if credentials password match with the one of given user.
        if (!cre.getPassword().equals(user.getPassword()))
            return null;

        // Session credentials are valid.
        return cre;
    }

    public boolean login(HttpSession session, String username, String password)
    {
        // Get the user with the given username.
        User user = userService.getUserByUsername(username);
        if (user == null)
            return false;

        // Validate the password.
        if (!user.getPassword().equals(password))
            return false;

        // Store a new user credential in session.
        UserCredential cre = new UserCredential(user);
        session.setAttribute("userCredential", cre);

        return true;
    }

    public void logout(HttpSession session)
    {
        session.removeAttribute("userCredential");
    }
}
