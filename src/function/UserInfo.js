import { URL_ROOT,ACCESS_TOKEN } from "../valueConst.js"

export function getUserInfo(){
    fetch(URL_ROOT+"userInfo",{
        method:"GET",
        redirect: "follow", // manual, *follow, error
        mode: "cors", // no-cors, *cors, same-origin
        cache: "default", // *default, no-cache, reload, force-cache, only-if-cached
        credentials: "same-origin",
        headers: {
            authorization: `${ACCESS_TOKEN}`,
            "X-Auth-Token":   `${ACCESS_TOKEN}`,
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Credentials": true,
          },
    })
    .then(function(respond){
        return respond.json()
    })
    .then(function(data){
        console.log(data)
    })
    .catch(function(error){
        console.log(error)
    })
}

