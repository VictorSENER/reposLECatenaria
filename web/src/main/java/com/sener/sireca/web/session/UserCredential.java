/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.session;

import java.io.Serializable;
import java.util.HashSet;
import java.util.Set;

import com.sener.sireca.web.bean.User;

public class UserCredential implements Serializable
{
    private static final long serialVersionUID = 1L;

    String username;
    String password;
    Set<String> roles = new HashSet<String>();

    public UserCredential(User u)
    {
        this.username = u.getUsername();
        this.password = u.getPassword();
    }

    public String getUsername()
    {
        return username;
    }

    public void setUsername(String username)
    {
        this.username = username;
    }

    public String getPassword()
    {
        return password;
    }

    public void setPassword(String password)
    {
        this.password = password;
    }

    public boolean hasRole(String role)
    {
        return roles.contains(role);
    }

    public void addRole(String role)
    {
        roles.add(role);
    }
}
