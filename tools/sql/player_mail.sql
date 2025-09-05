CREATE TABLE player_emails (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sender_id INT NOT NULL,
  target_id INT NOT NULL,
  email_id INT NOT NULL, -- Identificador comum para agrupar todos os itens de um único envio

  item_object_id INT NOT NULL,
  item_id SMALLINT NOT NULL,
  count INT NOT NULL,
  enchant_level SMALLINT NOT NULL,
  
  is_augmented BOOLEAN NOT NULL DEFAULT FALSE,
  augment_id INT DEFAULT NULL, -- Se for augmentado, guardar o ID do augment

  is_paid BOOLEAN NOT NULL DEFAULT FALSE,
  payment_item_id SMALLINT DEFAULT NULL, -- Ex: 57 = Adena, 3470 = Gold Bar, etc.
  payment_item_count INT DEFAULT NULL,

  status ENUM('PENDING', 'ACCEPTED', 'REJECTED', 'EXPIRED') DEFAULT 'PENDING',
  expiration_time BIGINT NOT NULL,
  created_time BIGINT NOT NULL
);
