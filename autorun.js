const BREE_CONFIG = {

    logoUrl:
        "https://bree.com.br/assinatura-email.png",

    // Futuro endpoint de producao.
    // Ainda vamos construir essa API.
    userApiUrl:
        "https://assinatura-api.bree.com.br/user",

    // Somente desenvolvimento local.
    localUsersUrl:
        "./users.json"
};


function clean(value) {

    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    return String(value).trim();
}


function escapeHtml(value) {

    return clean(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


function formatPhone(value) {

    const phone = clean(value);

    if (!phone) {
        return "";
    }

    if (phone.startsWith("+")) {
        return phone;
    }

    return "+55 " + phone;
}


function formatAddress(user) {

    const street =
        clean(user.street);

    const city =
        clean(user.city);

    const state =
        clean(user.state);

    const postalCode =
        clean(user.postalCode);

    let cityState = "";

    if (city && state) {
        cityState =
            city + "/" + state;
    }
    else {
        cityState =
            city || state;
    }

    const parts = [];

    if (street) {
        parts.push(street);
    }

    if (cityState) {
        parts.push(cityState);
    }

    let address =
        parts.join(", ");

    if (address) {
        address += " – Brazil";
    }

    if (postalCode) {

        if (address) {
            address += " ";
        }

        address += postalCode;
    }

    return address;
}


function isLocalDevelopment() {

    return (
        window.location.hostname === "localhost" ||
        window.location.hostname === "127.0.0.1"
    );
}


async function loadUser(email) {

    /*
     * Desenvolvimento:
     * carrega users.json local.
     */
    if (isLocalDevelopment()) {

        const response =
            await fetch(
                BREE_CONFIG.localUsersUrl +
                "?t=" +
                Date.now(),
                {
                    cache: "no-store"
                }
            );

        if (!response.ok) {

            throw new Error(
                "Falha ao carregar users.json. HTTP " +
                response.status
            );
        }

        const users =
            await response.json();

        return users[email] || null;
    }


    /*
     * Producao:
     * consulta somente o usuario atual.
     */
    const url =
        BREE_CONFIG.userApiUrl +
        "?email=" +
        encodeURIComponent(email);

    const response =
        await fetch(
            url,
            {
                method: "GET",
                cache: "no-store"
            }
        );

    if (response.status === 404) {
        return null;
    }

    if (!response.ok) {

        throw new Error(
            "Falha ao consultar usuario. HTTP " +
            response.status
        );
    }

    return await response.json();
}


function getSenderEmail(callback) {

    const item =
        Office.context.mailbox.item;

    /*
     * Tenta pegar o endereco atualmente
     * selecionado no campo De.
     */
    if (
        item.from &&
        item.from.getAsync
    ) {

        item.from.getAsync(
            function (result) {

                if (
                    result.status ===
                        Office.AsyncResultStatus.Succeeded &&
                    result.value &&
                    result.value.emailAddress
                ) {

                    callback(
                        result.value.emailAddress
                            .toLowerCase()
                            .trim()
                    );

                    return;
                }

                /*
                 * Fallback:
                 * usuario logado.
                 */
                callback(
                    Office.context.mailbox
                        .userProfile
                        .emailAddress
                        .toLowerCase()
                        .trim()
                );
            }
        );

        return;
    }


    callback(
        Office.context.mailbox
            .userProfile
            .emailAddress
            .toLowerCase()
            .trim()
    );
}


function buildSignature(user, email) {

    const displayName =
        escapeHtml(user.displayName);

    const titlePt =
        escapeHtml(user.titlePt);

    const titleEn =
        escapeHtml(user.titleEn);

    const extraLine =
        escapeHtml(user.extraLine);

    const phone =
        escapeHtml(
            formatPhone(user.phone)
        );

    const mobile =
        escapeHtml(
            formatPhone(user.mobile)
        );

    const address =
        escapeHtml(
            formatAddress(user)
        );

    const safeEmail =
        escapeHtml(email);


    let titleHtml = "";

    if (titlePt && titleEn) {

        titleHtml =
            titlePt +
            " / " +
            titleEn;
    }
    else {

        titleHtml =
            titlePt ||
            titleEn;
    }


    return `
<table
    cellpadding="0"
    cellspacing="0"
    border="0"
    style="
        font-family:Arial,Helvetica,sans-serif;
        font-size:14px;
        line-height:21px;
        color:#294f84;
    ">

    <tr>
        <td style="padding-bottom:10px;">
            Atenciosamente,
        </td>
    </tr>

    <tr>
        <td style="font-weight:bold;">
            ${displayName}
        </td>
    </tr>

    ${
        titleHtml
        ? `
        <tr>
            <td>
                ${titleHtml}
            </td>
        </tr>
        `
        : ""
    }

    ${
        extraLine
        ? `
        <tr>
            <td>
                ${extraLine}
            </td>
        </tr>
        `
        : ""
    }

    ${
        address
        ? `
        <tr>
            <td>
                ${address}
            </td>
        </tr>
        `
        : ""
    }

    ${
        phone
        ? `
        <tr>
            <td>
                ${phone}
            </td>
        </tr>
        `
        : ""
    }

    ${
        mobile
        ? `
        <tr>
            <td>
                ${mobile}
            </td>
        </tr>
        `
        : ""
    }

    <tr>
        <td>

            <a
                href="mailto:${safeEmail}"
                style="
                    color:#294f84;
                    text-decoration:underline;
                ">
                ${safeEmail}
            </a>

            /

            <a
                href="https://www.bree.com.br"
                style="
                    color:#7030A0;
                    text-decoration:underline;
                ">
                www.bree.com.br
            </a>

        </td>
    </tr>

    <tr>
        <td style="padding-top:10px;">

            <img
                src="${BREE_CONFIG.logoUrl}"
                alt="Bree"
                width="480"
                style="
                    display:block;
                    border:0;
                ">

        </td>
    </tr>

</table>
`;
}


function applyBreeSignature(event) {

    const host =
        Office.context.mailbox
            .diagnostics
            .hostName;

    /*
     * Outlook Classico:
     * tratado pela GPO.
     */
    if (
        host !== "newOutlookWindows" &&
        host !== "OutlookWebApp"
    ) {

        event.completed();
        return;
    }


    getSenderEmail(
        async function (email) {

            try {

                const user =
                    await loadUser(email);

                if (!user) {

                    console.error(
                        "Bree Signature: usuario nao encontrado:",
                        email
                    );

                    event.completed();
                    return;
                }


                const signature =
                    buildSignature(
                        user,
                        email
                    );


                Office.context.mailbox
                    .item
                    .body
                    .setSignatureAsync(

                        signature,

                        {
                            coercionType:
                                Office.CoercionType.Html,

                            asyncContext:
                                event
                        },

                        function (result) {

                            if (
                                result.status !==
                                Office.AsyncResultStatus.Succeeded
                            ) {

                                console.error(
                                    "Bree Signature:",
                                    result.error
                                );
                            }

                            result
                                .asyncContext
                                .completed();
                        }
                    );
            }
            catch (error) {

                console.error(
                    "Bree Signature:",
                    error
                );

                event.completed();
            }
        }
    );
}


/*
 * Novo email / resposta / encaminhamento.
 */
function onNewMessageComposeHandler(event) {

    applyBreeSignature(event);
}


/*
 * Usuario alterou o campo "De".
 */
function onMessageFromChangedHandler(event) {

    applyBreeSignature(event);
}


/*
 * Associacao exigida para event-based activation.
 */
Office.actions.associate(
    "onNewMessageComposeHandler",
    onNewMessageComposeHandler
);

Office.actions.associate(
    "onMessageFromChangedHandler",
    onMessageFromChangedHandler
);