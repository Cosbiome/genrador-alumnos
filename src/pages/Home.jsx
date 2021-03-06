import React, { useState } from "react";
import { remote } from "electron";
import excelToJson from "convert-excel-to-json";
import json2xls from "json2xls";
import { writeFileSync } from "fs";
import moment from "moment";

const { dialog, Notification } = remote;

const Home = () => {
  const [usuarios, setUsuarios] = useState([]);

  const capitalize = (word) => {
    return word.substr(0, 1).toUpperCase() + word.slice(1);
  };

  const handelFormatUsersFC = (data) => {
    return data["Hoja1"].map((alum) => {
      let facial = "";
      let corporal = "";
      let dia = moment().format("L").split("/")[1];

      let password = alum.NOMBRE.split(" ")
        .map((a) => capitalize(a.substr(0, 3).toLowerCase()))
        .join("");

      let nombre = alum.NOMBRE.toLowerCase()
        .replace(/\b\w/g, (l) => l.toUpperCase())
        .trim();

      let correoSplit = nombre.toLowerCase().split(" ");
      let correo = `${correoSplit[0]}${
        correoSplit[correoSplit.length - 1]
      }${dia}@cosbiome.mx`;
      let usuario = `${correoSplit[1]}.${
        correoSplit[correoSplit.length - 1]
      }${dia}`;

      if (alum["CORPORAL MARTES"] !== undefined) {
        corporal = "X";
      }
      if (alum["FACIAL LUNES"] !== undefined) {
        facial = "X";
      }

      return {
        nombre,
        usuario,
        password: password + dia,
        correo,
        facial,
        corporal,
      };
    });
  };

  const handleGetImagenes = () => {
    let file = dialog.showOpenDialogSync({
      properties: ["openFile"],
    });

    const excelConvert = excelToJson({
      sourceFile: file[0],
      columnToKey: {
        A: "NOMBRE",
        B: "TELEFONO",
        C: "MODALIDAD",
        D: "FACIAL LUNES",
        E: "CORPORAL MARTES",
      },
      header: {
        rows: 1,
      },
      sheets: ["Hoja1"],
    });

    let usuariosGEN = handelFormatUsersFC(excelConvert);

    setUsuarios(usuariosGEN);
  };

  const handleSaveDocument = (filt) => {
    let usuariosToExcel = usuarios;

    if (filt !== "") {
      usuariosToExcel = usuarios.filter((a) => a[filt] === "X");
    }

    let xls = json2xls(usuariosToExcel);

    dialog
      .showSaveDialog(null, {
        title: "USUARIOS COSBIOME",
        defaultPath: "Documents",
        buttonLabel: "guardar",
        filters: [{ name: "usuarios", extensions: ["xlsx"] }],
      })
      .then(({ filePath }) => {
        try {
          writeFileSync(filePath, xls, "binary");
          new Notification({
            title: "ARCHIVO GUARDADO",
            body: "EL ARCHIVO SE A GUARDADO EXITOSA MENTE EN " + filePath,
          }).show();
        } catch (error) {}
      });
  };

  return (
    <div className="container">
      <h1 className="text-center">GENERADOR DE ALUMNOS CORPORAL Y FACIAL</h1>
      <div className="row mt-5 d-flex justify-content-center">
        <div className="col-md-3">
          <button
            onClick={() => handleGetImagenes()}
            className="btn btn-primary"
          >
            INGRESAR EXCEL DE ALUMNOS
          </button>
        </div>
        {usuarios.length > 0 && (
          <>
            <div className="col-md-3">
              <button
                onClick={() => handleSaveDocument("facial")}
                className="btn btn-primary"
              >
                EXPORTAR FACIAL
              </button>
            </div>
            <div className="col-md-3">
              <button
                onClick={() => handleSaveDocument("corporal")}
                className="btn btn-primary"
              >
                EXPORTAR CORPORAL
              </button>
            </div>
            <div className="col-md-3">
              <button
                onClick={() => handleSaveDocument("")}
                className="btn btn-primary"
              >
                EXPORTAR JUNTOS
              </button>
            </div>
          </>
        )}
      </div>

      <div className="row mt-5">
        <div className="col-md-12">
          <table className="table text-center table-bordered">
            <thead>
              <tr>
                <th scope="col">NOMBRE</th>
                <th scope="col">USUARIO</th>
                <th scope="col">PASSWORD</th>
                <th scope="col">CORREO</th>
                <th scope="col">FACIAL</th>
                <th scope="col">CORPORAL</th>
              </tr>
            </thead>
            <tbody>
              {usuarios.map((usuario) => {
                return (
                  <tr key={usuario.password}>
                    <td> {usuario.nombre} </td>
                    <td> {usuario.usuario} </td>
                    <td> {usuario.password} </td>
                    <td> {usuario.correo} </td>
                    <td> {usuario.facial} </td>
                    <td> {usuario.corporal} </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default Home;
