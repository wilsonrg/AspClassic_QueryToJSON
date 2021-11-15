//========== Conteúdo ==========
/*Variáveis globais*/
let v1, v2, v3, v4;
$(function () {
    $.getJSON('json/json_query_login.asp', function (data) {
        v1 = ''; v2 = ''; v3 = ''; v4 = '';
        if (data.length > 0) {
            $.each(data, function (w, itens) {
                v1 = data[w].id; v1 = fv(v1);
                v2 = data[w].login; v2 = fv(v2);
                v3 = data[w].senha; v3 = fv(v3);
                v4 += '<p>Usuário: ' + v2 + ' - Senha: ' + v3 + ' </p>';
            });
            $('#conteudo').html(v4);
        }
    });
});
function fv(v) {
    const valor = v;
    return valor === undefined || valor === null ? '' : valor;
}