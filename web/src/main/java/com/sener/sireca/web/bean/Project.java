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
@Table(name = "proyecto")
public class Project
{
    // Identificador del usuario
    @Id
    @GeneratedValue(generator = "increment")
    @GenericGenerator(name = "increment", strategy = "increment")
    private Integer id;

    // Titulo del proyecto
    @Column(name = "titulo")
    private String titulo;

    // Id del usuario
    @Column(name = "idUsuario")
    private int idUsuario;

    // Id del cliente
    @Column(name = "cliente")
    private String cliente;

    // Referencia del Proyecto
    @Column(name = "referencia")
    private String referencia;

    public Integer getId()
    {
        return id;
    }

    public String getTitulo()
    {
        return titulo;
    }

    public void setTitulo(String titulo)
    {
        this.titulo = titulo;
    }

    public int getIdUsuario()
    {
        return idUsuario;
    }

    public void setIdUsuario(int idUsuario)
    {
        this.idUsuario = idUsuario;
    }

    public String getCliente()
    {
        return cliente;
    }

    public void setCliente(String cliente)
    {
        this.cliente = cliente;
    }

    public String getReferencia()
    {
        return referencia;
    }

    public void setReferencia(String referencia)
    {
        this.referencia = referencia;
    }

    private String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + id + Globals.CALCULO_REPLANTEO;
    }

    public String getCalcReplanteoBasePath()
    {
        return getBasePath() + Globals.CALCULO_REPLANTEO;
    }

    public String getDibReplanteoBasePath()
    {
        return getBasePath() + Globals.DIBUJO_REPLANTEO;
    }

    public String getMonReplanteoBasePath()
    {
        return getBasePath() + Globals.FICHAS_MONTAJE;
    }

    public String getPenReplanteoBasePath()
    {
        return getBasePath() + Globals.FICHAS_PENDOLADO;
    }

}
