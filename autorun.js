Office.onReady();

function clean(value) {
    if (value === null || value === undefined) {
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

    // Se ja possui codigo internacional, mantem.
    if (phone.startsWith("+")) {
        return phone;
    }

    // Telefones brasileiros armazenados como (41) 3167-XXXX.
    return "+55 " + phone;
}

function formatAddress(user) {

    const street = clean(user.street);
    const city = clean(user.city);
    const state = clean(user.state);
    const postalCode = clean(user.postalCode);

    let location = "";

    if (city && state) {
        location = city + "/" + state;
    }
    else if (city) {
        location = city;
    }
    else if (state) {
        location = state;
    }

    const parts = [];

    if (street) {
        parts.push(street);
    }

    if (location) {
        parts.push(location);
    }

    let address = parts.join(", ");

    if (address) {
        address += " – Brazil";
    }

    if (postalCode) {
        address += (address ? " " : "") + postalCode;
    }

    return address;
}

function getSenderEmail(callback) {

    const item = Office.context.mailbox.item;

    // Pega o endereco que esta efetivamente no campo "De".
    if (item.from && item.from.getAsync) {

        item.from.getAsync(function (result) {

            if (
                result.status === Office.AsyncResultStatus.Succeeded &&
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

            // Fallback para a conta principal.
            callback(
                Office.context.mailbox.userProfile.emailAddress
                    .toLowerCase()
                    .trim()
            );
        });

        return;
    }

    callback(
        Office.context.mailbox.userProfile.emailAddress
            .toLowerCase()
            .trim()
    );
}

async function loadUsers() {

    const response = await fetch(
        "https://localhost:3000/users.json?t=" + Date.now(),
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

    return await response.json();
}

function buildSignature(user, email) {

    const displayName = escapeHtml(user.displayName);

    const titlePt = escapeHtml(user.titlePt);
    const titleEn = escapeHtml(user.titleEn);

    const extraLine = escapeHtml(user.extraLine);

    const phone = escapeHtml(
        formatPhone(user.phone)
    );

    const mobile = escapeHtml(
        formatPhone(user.mobile)
    );

    const address = escapeHtml(
        formatAddress(user)
    );

    const safeEmail = escapeHtml(email);

    let titleHtml = "";

    if (titlePt && titleEn) {
        titleHtml =
            titlePt + " / " + titleEn;
    }
    else {
        titleHtml =
            titlePt || titleEn;
    }

    let extraHtml = "";

    if (extraLine) {
        extraHtml = `
            <tr>
                <td>
                    ${extraLine}
                </td>
            </tr>
        `;
    }

    let phoneHtml = "";

    if (phone) {
        phoneHtml = `
            <tr>
                <td>
                    ${phone}
                </td>
            </tr>
        `;
    }

    let mobileHtml = "";

    if (mobile) {
        mobileHtml = `
            <tr>
                <td>
                    ${mobile}
                </td>
            </tr>
        `;
    }

    let addressHtml = "";

    if (address) {
        addressHtml = `
            <tr>
                <td>
                    ${address}
                </td>
            </tr>
        `;
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
        <td style="padding:0 0 10px 0;">
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

    ${extraHtml}

    ${addressHtml}

    ${phoneHtml}

    ${mobileHtml}

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
                src="https://bree.com.br/assinatura-email.png"
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
        Office.context.mailbox.diagnostics.hostName;

    // O Outlook Classico sera tratado pela GPO.
    if (
        host !== "newOutlookWindows" &&
        host !== "OutlookWebApp"
    ) {
        event.completed();
        return;
    }

    getSenderEmail(async function (email) {

        try {

            const users = await loadUsers();

            const user = users[email];

            if (!user) {

                console.error(
                    "Bree Signature: usuario nao encontrado:",
                    email
                );

                event.completed();
                return;
            }

            const signature =
                buildSignature(user, email);

            Office.context.mailbox.item.body
                .setSignatureAsync(
                    signature,
                    {
                        coercionType:
                            Office.CoercionType.Html,
                        asyncContext: event
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

                        result.asyncContext.completed();
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
    });
}

function onMessageComposeHandler(event) {
    applyBreeSignature(event);
}

Office.actions.associate(
    "onMessageComposeHandler",
    onMessageComposeHandler
);