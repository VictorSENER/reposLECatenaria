/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;

import org.hibernate.annotations.GenericGenerator;

@Entity
@Table(name = "Datos")
public class Catenaria
{
    // Identificador del usuario
    @Id
    @GeneratedValue(generator = "increment")
    @GenericGenerator(name = "increment", strategy = "increment")
    private Integer id;

    // Titulo del proyecto
    @Column(name = "nombre_cat")
    private String nomCatenaria;

    public Integer getId()
    {
        return id;
    }

    public String getNomCatenaria()
    {
        return nomCatenaria;
    }
}
