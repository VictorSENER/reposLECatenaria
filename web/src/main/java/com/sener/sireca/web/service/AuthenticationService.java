/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import javax.servlet.http.HttpSession;

import com.sener.sireca.web.session.UserCredential;

public interface AuthenticationService
{

    public UserCredential getUserCredential(HttpSession session);

    public boolean login(HttpSession session, String account, String password);

    public void logout(HttpSession session);

}
