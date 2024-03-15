import pandas as pd
import json


# Ruta al archivo Excel
archivo_excel = "C:/Users/Edwar/Documents/proyectosPython/Sofiia-RIPs-/venv/Excel4Rips.xlsx"

# Leer el archivo Excel
datos_excel = pd.read_excel(archivo_excel)

# Mostrar los datos
print ("imprimiendo...")
print(datos_excel)


json_data = {

    "numDocumentoIdObligado": str(datos_excel["numDocumentoIdObligado"][0]),

    "numFactura": str(datos_excel["numFactura"][0]),

    "TipoNota": str(datos_excel["TipoNota"][0]),

    "numNota": str(datos_excel["numNota"][0]),

    "usuarios": [

        {

            "tipoDocumentoIdentificacion": str(datos_excel["tipoDocumentoIdentificacion"][0]),

            "numDocumentoIdentificacion": str(datos_excel["numDocumentoIdentificacion"][0]),

            "tipoUsuario": str(datos_excel["tipoUsuario"][0]),

            "fechaNacimiento": str(datos_excel["fechaNacimiento"][0]),

            "codSexo": str(datos_excel["codSexo"][0]),

            "codPaisResidencia": str(datos_excel["codPaisResidencia"][0]),

            "codMunicipioResidencia": str(datos_excel["codMunicipioResidencia"][0]),

            "codZonaTerritorialResidencia": str(datos_excel["codZonaTerritorialResidencia"][0]),

            "incapacidad": str(datos_excel["incapacidad"][0]),

            "consecutivo": str(datos_excel["consecutivo"][0]),

            "codPaisOrigen": str(datos_excel["codPaisOrigen"][0]),

            "servicios": {

                "consultas": [

                    {

                        "codPrestador": "",

                        "fechaInicioAtencion": "",

                        "numAutorizacion": "",

                        "codConsulta": "",

                        "modalidadGrupoServicioTecSal": "",

                        "grupoServicios": "",

                        "codServicio": "",

                        "finalidadTecnologiaSalud": "",

                        "causaMotivoAtencion": "",

                        "codDiagnosticoPrincipal": "",

                        "codDiagnosticoRelacionado1": "",

                        "codDiagnosticoRelacionado2": "null",

                        "codDiagnosticoRelacionado3": "null",

                        "tipoDiagnosticoPrincipal": "",

                        "tipoDocumentoIdentificacion": "",

                        "numDocumentoIdentificacion": "",

                        "vrServicio": "",

                        "conceptoRecaudo": "",

                        "valorPagoModerador": "",

                        "numFEVPagoModerador": "",

                        "consecutivo": ""

                    },

                    {

                        "codPrestador": "",

                        "fechaInicioAtencion": "",

                        "numAutorizacion": "",

                        "codConsulta": "",

                        "modalidadGrupoServicioTecSal": "",

                        "grupoServicios": "",

                        "codServicio": "",

                        "finalidadTecnologiaSalud": "",

                        "causaMotivoAtencion": "",

                        "codDiagnosticoPrincipal": "",

                        "codDiagnosticoRelacionado1": "",

                        "codDiagnosticoRelacionado2": "null",

                        "codDiagnosticoRelacionado3": "null",

                        "tipoDiagnosticoPrincipal": "",

                        "tipoDocumentoIdentificacion": "",

                        "numDocumentoIdentificacion": "",

                        "vrServicio": "",

                        "conceptoRecaudo": "",

                        "valorPagoModerador": "",

                        "numFEVPagoModerador": "",

                        "consecutivo": ""

                    },

                ],

                "procedimientos": [

                    {

                        "codPrestador": "",

                        "fechaInicioAtencion": "",

                        "idMIPRES": "null",

                        "numAutorizacion": "null",

                        "codProcedimiento": "",

                        "viaIngresoServicioSalud": "",

                        "modalidadGrupoServicioTecSal": "",

                        "grupoServicios": "",

                        "codServicio": "",

                        "finalidadTecnologiaSalud": "",

                        "tipoDocumentoIdentificacion": "",

                        "numDocumentoIdentificacion": "",

                        "codDiagnosticoPrincipal": "",

                        "codDiagnosticoRelacionado": "",

                        "codComplicacion": "",

                        "vrServicio": "",

                        "conceptoRecaudo": "",

                        "valorPagoModerador": "",

                        "numFEVPagoModerador": "",

                        "consecutivo": ""

                    },

                ],

                "urgencias": [

                    {

                        "codPrestador": "",

                        "fechaInicioAtencion": "",

                        "causaMotivoAtencion": "",

                        "codDiagnosticoPrincipal": "",

                        "codDiagnosticoPrincipalE": "",

                        "codDiagnosticoRelacionadoE1": "null",

                        "codDiagnosticoRelacionadoE2": "null",

                        "codDiagnosticoRelacionadoE3": "null",

                        "condicionDestino": "",

                        "codDiagnosticoCausaMuerte": "null",

                        "fechaEgreso": "",

                        "consecutivo": ""

                    },

                ],

                "hospitalizacion": [

                    {

                        "codPrestador": "",

                        "viaIngresoServicioSalud": "",

                        "fechaInicioAtencion": "",

                        "numAutorizacion": "",

                        "causaMotivoAtencion": "",

                        "codDiagnosticoPrincipal": "",

                        "codDiagnosticoPrincipalE": "",

                        "codDiagnosticoRelacionadoE1": "null",

                        "codDiagnosticoRelacionadoE2": "null",

                        "codDiagnosticoRelacionadoE3": "null",

                        "codComplicacion": "null",

                        "condicionDestinoUsuarioEgreso": "",

                        "codDiagnosticoCausaMuerte": "null",

                        "fechaEgreso": "",

                        "consecutivo": ""

                    },

                ],

                "recienNacidos": [

                    {

                        "codPrestador": "",

                        "tipoDocumentoIdentificacion": "",

                        "numDocumentoIdentificacion": "",

                        "fechaNacimiento": "",

                        "edadGestacional": "",

                        "numConsultasCPrenatal": "",
    
                        "codSexoBiologico": "",

                        "peso": "",

                        "codDiagnosticoPrincipal": "",

                        "condicionDestinoUsuarioEgreso": "",

                        "codDiagnosticoCausaMuerte": "",

                        "fechaEgreso": "",

                        "consecutivo": ""

                    },

                ],

                "medicamentos": [

                    {

                        "codPrestador": "",

                        "numAutorizacion": "",

                        "idMIPRES": "",

                        "fechaDispensAdmon": "",

                        "codDiagnosticoPrincipal": "",

                        "codDiagnosticoRelacionado": "",

                        "tipoMedicamento": "",

                        "codTecnologiaSalud": "",

                        "nomTecnologiaSalud": "null",

                        "concentracionMedicamento": "",

                        "unidadMedida": "",

                        "formaFarmaceutica": "",

                        "unidadMinDispensa": "",

                        "cantidadMedicamento": "",

                        "diasTratamiento": "",

                        "tipoDocumentoIdentificacion": "",

                        "numDocumentoIdentificacion": "",

                        "vrUnitMedicamento": "",

                        "vrServicio": "",

                        "conceptoRecaudo": "",

                        "valorPagoModerador": "",

                        "numFEVPagoModerador": "",

                        "consecutivo": ""

                    },

                ],

                "otrosServicios": [

                    {

                        "codPrestador": str(datos_excel["codPrestador"][0]),

                        "numAutorizacion": str(datos_excel["numAutorizacion"][0]),

                        "idMIPRES": str(datos_excel["idMIPRES"][0]),

                        "fechaSuministroTecnologia": str(datos_excel["fechaSuministroTecnologia"][0]),

                        "tipoOS": str(datos_excel["tipoOS"][0]),

                        "codTecnologiaSalud": str(datos_excel["codTecnologiaSalud"][0]),

                        "nomTecnologiaSalud": str(datos_excel["nomTecnologiaSalud"][0]),

                        "cantidadOS": str(datos_excel["cantidadOS"][0]),

                        "tipoDocumentoIdentificacion": str(datos_excel["tipoDocumentoIdentificacionOtrosServicios"][0]),

                        "numDocumentoIdentificacion": str(datos_excel["numDocumentoIdentificacionOtrosServicios"][0]),

                        "vrUnitOS": str(datos_excel["vrUnitOS"][0]),

                        "vrServicio": str(datos_excel["vrServicio"][0]),

                        "conceptoRecaudo": str(datos_excel["conceptoRecaudo"][0]),

                        "valorPagoModerador": str(datos_excel["valorPagoModerador"][0]),

                        "numFEVPagoModerador": str(datos_excel["numFEVPagoModerador"][0]),

                        "consecutivo": str(datos_excel["consecutivoOtrosServicios"][0])

                    },
                ],

            },

        },

    ]

}


# Specify the file path where you want to save the JSON file
file_path = "C:/Users/Edwar/Documents/proyectosPython/Sofiia-RIPs-/venv/example.json"

# Write the data to the JSON file
with open(file_path, 'w') as json_file:
    json.dump(json_data, json_file, indent=4)

print(f"JSON data has been saved to {file_path}")