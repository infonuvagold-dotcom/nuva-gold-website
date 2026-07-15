(function () {
  function getByPath(obj, path) {
    return path.split('.').reduce(function (acc, key) {
      return acc == null ? undefined : acc[key];
    }, obj);
  }

  function digits(str) {
    return (str || '').replace(/\D/g, '');
  }

  function applyContent(data) {
    document.querySelectorAll('[data-key]').forEach(function (el) {
      var key = el.getAttribute('data-key');
      var val = getByPath(data, key);
      if (val === undefined || val === null) return;
      el.textContent = val;
    });

    if (data.contact) {
      var wa1 = document.getElementById('waBtn1');
      var wa2 = document.getElementById('waBtn2');
      var mailBtn = document.getElementById('mailBtn');

      if (wa1 && data.contact.phone1) {
        wa1.href = 'https://wa.me/' + digits(data.contact.phone1);
        wa1.textContent = 'WhatsApp — ' + data.contact.phone1.replace('+234 ', '0');
      }
      if (wa2 && data.contact.phone2) {
        wa2.href = 'https://wa.me/' + digits(data.contact.phone2);
        wa2.textContent = 'WhatsApp — ' + data.contact.phone2.replace('+234 ', '0');
      }
      if (mailBtn && data.contact.email) {
        mailBtn.href = 'mailto:' + data.contact.email;
      }
    }
  }

  fetch('content/site.json', { cache: 'no-store' })
    .then(function (res) {
      if (!res.ok) throw new Error('content/site.json not found');
      return res.json();
    })
    .then(applyContent)
    .catch(function (err) {
      // Falls back silently to the text already baked into the HTML
      console.warn('NUVA-Gold content loader:', err.message);
    });
})();
