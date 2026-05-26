
let cars = [];
        let reservations = [];
        let currentTab = 'cars';
        let editingCarId = null;
        let deferredPrompt = null;
        let currentCarFilter = 'all';
        let currentResFilter = 'all';

        function initializeApp() {
            setTimeout(() => { document.getElementById('splash').classList.add('hidden'); }, 2000);
            document.getElementById('settings-url').value = CONFIG.API_URL;
            document.getElementById('settings-secret').value = CONFIG.PWA_SECRET;
            loadCars();
            loadReservations();
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('res-start').value = today;
            document.getElementById('res-end').value = today;
            window.addEventListener('beforeinstallprompt', (e) => {
                e.preventDefault();
                deferredPrompt = e;
                document.getElementById('install-prompt').classList.add('active');
            });
            let touchStartY = 0;
            document.addEventListener('touchstart', (e) => { touchStartY = e.touches[0].clientY; });
            document.addEventListener('touchend', (e) => {
                const touchEndY = e.changedTouches[0].clientY;
                if (touchEndY - touchStartY > 150 && window.scrollY === 0) refreshData();
            });
        }
        async function loadCars() {
            try {
                const data = await apiCall('list');
                if (data.ok && data.cars) {
                    cars = data.cars;
                    renderCars();
                    updateStats();
                }
            } catch (e) { console.error('Erreur chargement voitures:', e); }
        }

        async function loadReservations() {
            try {
                const data = await apiCall('listReservations');
                reservations = data.reservations || [];
                renderReservations();
                updateStats();
            } catch (e) { console.error('Erreur chargement reservations:', e); }
        }

        function renderCars(filter = currentCarFilter, search = document.getElementById('car-search').value) {
            const container = document.getElementById('cars-list');
            let filtered = cars;
            if (filter !== 'all') {
                if (filter === 'automatic' || filter === 'manual') {
                    filtered = filtered.filter(c => c.transmission.toLowerCase() === filter);
                } else {
                    filtered = filtered.filter(c => c.fuel.toLowerCase() === filter);
                }
            }
            if (search) {
                filtered = filtered.filter(c => c.name.toLowerCase().includes(search.toLowerCase()));
            }
            if (filtered.length === 0) {
                container.innerHTML = `<div class="empty-state"><i class="fas fa-car-side"></i><h3>Aucun vehicule</h3><p>${search ? 'Aucun resultat' : 'Ajoutez votre premier vehicule'}</p></div>`;
                return;
            }
            container.innerHTML = filtered.map(car => `
                <div class="card car-card" data-id="${car.id}">
                    <img src="${car.image}" alt="${car.name}" class="car-image" onerror="this.src='https://via.placeholder.com/100x75?text=Car'">
                    <div class="car-info">
                        <div class="car-name">${car.name}</div>
                        <div class="car-price">${car.price} MAD <span style="font-size:13px;font-weight:500;color:var(--secondary)">/ ${car.duration}</span></div>
                        <div class="car-meta">
                            <span class="car-tag"><i class="fas fa-user" style="font-size:9px;margin-right:3px"></i>${car.seats}</span>
                            <span class="car-tag"><i class="fas fa-cog" style="font-size:9px;margin-right:3px"></i>${car.transmission}</span>
                            <span class="car-tag"><i class="fas fa-gas-pump" style="font-size:9px;margin-right:3px"></i>${car.fuel}</span>
                            <span class="car-tag"><i class="fas fa-door-open" style="font-size:9px;margin-right:3px"></i>${car.doors}</span>
                        </div>
                    </div>
                    <div class="car-actions">
                        <button class="btn-edit" onclick="editCar(${car.id})" title="Modifier"><i class="fas fa-pen"></i></button>
                        <button class="btn-delete" onclick="deleteCar(${car.id})" title="Supprimer"><i class="fas fa-trash"></i></button>
                    </div>
                </div>
            `).join('');
        }

        function renderReservations(filter = currentResFilter, search = document.getElementById('res-search').value) {
            const container = document.getElementById('reservations-list');
            let filtered = reservations;
            if (filter !== 'all') filtered = filtered.filter(r => r.status === filter);
            if (search) {
                filtered = filtered.filter(r =>
                    (r.name && r.name.toLowerCase().includes(search.toLowerCase())) ||
                    (r.carName && r.carName.toLowerCase().includes(search.toLowerCase()))
                );
            }
            if (filtered.length === 0) {
                container.innerHTML = `<div class="empty-state"><i class="fas fa-calendar"></i><h3>Aucune reservation</h3><p>Les reservations apparaitront ici</p></div>`;
                return;
            }
            container.innerHTML = filtered.map(res => {
                const statusClass = getStatusClass(res.status);
                return `<div class="card res-card status-${statusClass}" data-id="${res.rowNumber || 0}">
                    <div class="res-header">
                        <div class="res-customer">${res.name || 'Sans nom'}</div>
                        <span class="res-status status-${statusClass}">${res.status || 'En attente'}</span>
                    </div>
                    <div class="res-car"><i class="fas fa-car"></i><span>${res.carName || 'Non specifie'} — ${res.carPrice || 0} MAD/${res.carDuration || 'jour'}</span></div>
                    <div class="res-details">
                        <div class="res-detail"><span class="label">Email</span><span class="value">${res.email || '-'}</span></div>
                        <div class="res-detail"><span class="label">Telephone</span><span class="value">${res.phone || '-'}</span></div>
                        <div class="res-detail"><span class="label">Debut</span><span class="value">${formatDate(res.startDate)}</span></div>
                        <div class="res-detail"><span class="label">Fin</span><span class="value">${formatDate(res.endDate)}</span></div>
                    </div>
                    <div class="res-actions">
                        ${res.status === 'En attente' ? `
                            <button class="btn-confirm" onclick="updateReservationStatus(${res.rowNumber}, 'Confirmee')"><i class="fas fa-check"></i> Confirmer</button>
                            <button class="btn-refuse" onclick="updateReservationStatus(${res.rowNumber}, 'Refusee')"><i class="fas fa-times"></i> Refuser</button>
                        ` : res.status === 'Confirmee' ? `
                            <button class="btn-confirm" onclick="updateReservationStatus(${res.rowNumber}, 'Terminee')"><i class="fas fa-check-double"></i> Terminer</button>
                            <button class="btn-cancel" onclick="updateReservationStatus(${res.rowNumber}, 'Annulee')"><i class="fas fa-ban"></i> Annuler</button>
                        ` : `<button class="btn-secondary" style="flex:1" onclick="showToast('Statut: ${res.status}', 'info')"><i class="fas fa-info-circle"></i> ${res.status}</button>`}
                    </div>
                </div>`;
            }).join('');
        }

        async function saveCar() {
            const car = {
                name: document.getElementById('car-name').value.trim(),
                price: parseFloat(document.getElementById('car-price').value) || 0,
                duration: document.getElementById('car-duration').value,
                seats: parseInt(document.getElementById('car-seats').value) || 4,
                transmission: document.getElementById('car-transmission').value,
                doors: parseInt(document.getElementById('car-doors').value) || 4,
                fuel: document.getElementById('car-fuel').value,
                image: document.getElementById('car-image').value.trim() || 'https://via.placeholder.com/100x75?text=Car'
            };
            if (!car.name || !car.price) {
                showToast('Veuillez remplir les champs obligatoires', 'warning');
                return;
            }
            try {
                if (editingCarId) {
                    car.id = editingCarId;
                    await apiPost('update', { car });
                    showToast('Vehicule modifie avec succes', 'success');
                } else {
                    await apiPost('add', { car });
                    showToast('Vehicule ajoute avec succes', 'success');
                }
                closeModal('car-modal');
                loadCars();
            } catch (e) { console.error(e); }
        }

        async function deleteCar(id) {
            if (!confirm('Supprimer ce vehicule ?')) return;
            try {
                await apiPost('delete', { id });
                showToast('Vehicule supprime', 'success');
                loadCars();
            } catch (e) { console.error(e); }
        }

        function editCar(id) {
            const car = cars.find(c => c.id === id);
            if (!car) return;
            editingCarId = id;
            document.getElementById('car-modal-title').textContent = 'Modifier vehicule';
            document.getElementById('car-name').value = car.name;
            document.getElementById('car-price').value = car.price;
            document.getElementById('car-duration').value = car.duration;
            document.getElementById('car-seats').value = car.seats;
            document.getElementById('car-transmission').value = car.transmission;
            document.getElementById('car-doors').value = car.doors;
            document.getElementById('car-fuel').value = car.fuel;
            document.getElementById('car-image').value = car.image;
            previewImage();
            openModal('car-modal');
        }

        async function saveReservation() {
            const reservation = {
                name: document.getElementById('res-name').value.trim(),
                email: document.getElementById('res-email').value.trim(),
                phone: document.getElementById('res-phone').value.trim(),
                carName: document.getElementById('res-car').value,
                startDate: document.getElementById('res-start').value,
                endDate: document.getElementById('res-end').value,
                notes: document.getElementById('res-notes').value.trim()
            };
            const car = cars.find(c => c.name === reservation.carName);
            if (car) { reservation.carPrice = car.price; reservation.carDuration = car.duration; }
            if (!reservation.name || !reservation.email || !reservation.phone || !reservation.carName || !reservation.startDate || !reservation.endDate) {
                showToast('Veuillez remplir tous les champs obligatoires', 'warning');
                return;
            }
            if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(reservation.email)) {
                showToast('Email invalide', 'warning');
                return;
            }
            if (new Date(reservation.endDate) < new Date(reservation.startDate)) {
                showToast('La date fin doit etre apres la date debut', 'warning');
                return;
            }
            try {
                await apiPost('createReservation', { reservation });
                showToast('Reservation creee avec succes', 'success');
                closeModal('res-modal');
                document.getElementById('res-name').value = '';
                document.getElementById('res-email').value = '';
                document.getElementById('res-phone').value = '';
                document.getElementById('res-notes').value = '';
                loadReservations();
            } catch (e) { console.error(e); }
        }

        async function updateReservationStatus(rowNumber, status) {
            if (!rowNumber) {
                showToast('Reservation introuvable', 'error');
                return;
            }
            try {
                await apiPost('updateReservationStatus', { rowNumber, status });
                showToast('Statut mis a jour: ' + status, 'success');
                loadReservations();
            } catch (e) { console.error(e); }
        }

        async function syncToGitHub() {
            try {
                await apiCall('sync');
                showToast('Synchronisation GitHub reussie', 'success');
                closeModal('sync-modal');
            } catch (e) { console.error(e); }
        }

        function switchTab(tab) {
            currentTab = tab;
            document.querySelectorAll('.content-area').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
            document.getElementById(tab + '-section').classList.add('active');
            document.querySelectorAll('.tab-btn').forEach((btn, idx) => {
                if ((tab === 'cars' && idx === 0) || (tab === 'reservations' && idx === 1)) btn.classList.add('active');
            });
            document.querySelectorAll('.nav-item').forEach((btn, idx) => {
                if ((tab === 'cars' && idx === 0) || (tab === 'reservations' && idx === 1)) btn.classList.add('active');
            });
            document.getElementById('fab-cars').style.display = tab === 'cars' ? 'flex' : 'none';
            document.getElementById('fab-res').style.display = tab === 'reservations' ? 'flex' : 'none';
        }

        function openCarModal() {
            editingCarId = null;
            document.getElementById('car-modal-title').textContent = 'Nouveau vehicule';
            document.getElementById('car-name').value = '';
            document.getElementById('car-price').value = '';
            document.getElementById('car-image').value = '';
            document.getElementById('car-image-preview').src = 'https://via.placeholder.com/400x160?text=Image+du+vehicule';
            openModal('car-modal');
        }

        function openReservationModal() {
            const select = document.getElementById('res-car');
            select.innerHTML = '<option value="">Choisir un vehicule...</option>' +
                cars.map(c => `<option value="${c.name}">${c.name} — ${c.price} MAD</option>`).join('');
            openModal('res-modal');
        }

        function openModal(id) {
            document.getElementById(id).classList.add('active');
            document.body.style.overflow = 'hidden';
        }

        function closeModal(id) {
            document.getElementById(id).classList.remove('active');
            document.body.style.overflow = '';
        }

        function showSyncModal() { openModal('sync-modal'); }
        function showSettings() { openModal('settings-modal'); }

        function saveSettings() {
            const url = document.getElementById('settings-url').value;
            const secret = document.getElementById('settings-secret').value;
            if (url) { localStorage.setItem('yaro_api_url', url); CONFIG.API_URL = url; }
            if (secret) { localStorage.setItem('yaro_secret', secret); CONFIG.PWA_SECRET = secret; }
            showToast('Parametres sauvegardes', 'success');
            closeModal('settings-modal');
        }

        function refreshData() {
            showToast('Rafraichissement...', 'info');
            loadCars();
            loadReservations();
        }

        function previewImage() {
            const url = document.getElementById('car-image').value;
            document.getElementById('car-image-preview').src = url || 'https://via.placeholder.com/400x160?text=Image+du+vehicule';
        }

        function filterCars() { renderCars(currentCarFilter); }
        function filterReservations() { renderReservations(currentResFilter); }

        function filterByType(type) {
            currentCarFilter = type;
            document.querySelectorAll('#car-filters .chip').forEach(c => c.classList.remove('active'));
            event.target.classList.add('active');
            renderCars(type);
        }

        function filterByStatus(status) {
            currentResFilter = status;
            document.querySelectorAll('#res-filters .chip').forEach(c => c.classList.remove('active'));
            event.target.classList.add('active');
            renderReservations(status);
        }

        function updateStats() {
            document.getElementById('stat-cars').textContent = cars.length;
            document.getElementById('tab-badge-cars').textContent = cars.length;
            document.getElementById('stat-res').textContent = reservations.length;
            document.getElementById('tab-badge-res').textContent = reservations.length;
            const pending = reservations.filter(r => r.status === 'En attente').length;
            document.getElementById('stat-pending').innerHTML = `<i class="fas fa-clock"></i> ${pending} en attente`;
        }

        function formatDate(value) {
            if (!value) return '-';
            if (value instanceof Date) return value.toLocaleDateString('fr-FR');
            const d = new Date(value);
            return isNaN(d.getTime()) ? String(value) : d.toLocaleDateString('fr-FR');
        }

        function getStatusClass(status) {
            return String(status || 'En attente')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/\s+/g, '-');
        }

        function showToast(message, type = 'info') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = `toast ${type}`;
            const icons = { success: 'fa-check-circle', error: 'fa-exclamation-circle', warning: 'fa-exclamation-triangle', info: 'fa-info-circle' };
            toast.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i><span>${message}</span>`;
            container.appendChild(toast);
            setTimeout(() => { toast.style.opacity = '0'; toast.style.transform = 'translateY(-20px)'; setTimeout(() => toast.remove(), 300); }, 3000);
        }

        function showLoading(show) {
            document.getElementById('loading').classList.toggle('active', show);
        }

        async function installApp() {
            if (deferredPrompt) {
                deferredPrompt.prompt();
                const { outcome } = await deferredPrompt.userChoice;
                if (outcome === 'accepted') showToast('Application installee !', 'success');
                deferredPrompt = null;
                document.getElementById('install-prompt').classList.remove('active');
            }
        }
